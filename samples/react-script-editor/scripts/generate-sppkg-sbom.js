'use strict';

const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

const root = path.resolve(__dirname, '..');
const evidenceDirectory = path.join(root, 'temp', 'sbom');
const hash = data => crypto.createHash('sha256').update(data).digest('hex');
const readJson = file => JSON.parse(fs.readFileSync(file, 'utf8'));
const fail = message => { throw new Error(message); };
const relative = file => path.relative(root, file).split(path.sep).join('/');

function packageForResource(resource, lock) {
    const file = resource.split('?')[0];
    if (!path.isAbsolute(file) || !fs.existsSync(file)) fail(`Missing module resource: ${resource}`);
    const rel = relative(file);
    if (rel.startsWith('../') || path.isAbsolute(rel)) fail(`Module outside project: ${resource}`);
    if (!rel.includes('node_modules/')) return undefined;
    let dir = path.dirname(file);
    while (dir !== root && dir !== path.dirname(dir)) {
        const manifestFile = path.join(dir, 'package.json');
        if (fs.existsSync(manifestFile)) {
            const manifest = readJson(manifestFile);
            const packagePath = relative(dir);
            const locked = lock.packages[packagePath];
            if (!locked || !manifest.name || !manifest.version || locked.version !== manifest.version) {
                fail(`Package does not match lockfile: ${packagePath}`);
            }
            const manifestPath = relative(manifestFile);
            return {
                ref: `${manifest.name}@${manifest.version}#${hash(Buffer.from(manifestPath))}`,
                path: packagePath,
                manifestSha256: hash(fs.readFileSync(manifestFile))
            };
        }
        dir = path.dirname(dir);
    }
    fail(`Missing package manifest for module: ${resource}`);
}

async function generate(packagePath, provenance) {
    const lockBytes = fs.readFileSync(path.join(root, 'package-lock.json'));
    const lock = JSON.parse(lockBytes.toString('utf8'));
    const capture = readJson(path.join(evidenceDirectory, 'modules.json'));
    const bom = readJson(path.join(evidenceDirectory, 'raw-bom.json'));
    const packageJson = readJson(path.join(root, 'package.json'));
    const config = readJson(path.join(root, 'config', 'package-solution.json'));
    const solution = config.solution;
    if (![2, 3].includes(lock.lockfileVersion) || capture.lockSha256 !== hash(lockBytes)) {
        fail('The lockfile is unsupported or differs from the captured build.');
    }
    if (solution.version !== `${packageJson.version}.0`) fail('Package and solution versions do not match.');

    const bytes = fs.readFileSync(packagePath);
    const archive = await JSZip.loadAsync(bytes, { checkCRC32: true });
    const manifestEntry = archive.file('AppManifest.xml');
    if (!manifestEntry) fail('SPPKG app manifest is missing.');
    const manifest = await manifestEntry.async('string');
    if (!manifest.includes(`Version="${solution.version}"`) ||
        !manifest.includes(`ProductID="${solution.id}"`)) fail('SPPKG identity or version differs from the project.');

    const files = new Map();
    const packages = new Map();
    const moduleEvidence = new Map();
    for (const chunk of capture.chunks) {
        if (!chunk.files.length) continue;
        for (const file of chunk.files) {
            if (files.has(file)) fail(`Ambiguous chunk asset: ${file}`);
            const entry = archive.file(`ClientSideAssets/${file}`);
            if (!entry || hash(await entry.async('nodebuffer')) !== chunk.hashes[file]) {
                fail(`SPPKG bundle does not match captured webpack asset: ${file}`);
            }
            files.set(file, chunk.hashes[file]);
        }
        for (const module of chunk.modules) {
            if (!module.resource) {
                if (module.kind === 'ExternalModule' || module.type === 'runtime') continue;
                if (module.kind === 'ConcatenatedModule' && module.nestedCount > 0) continue;
                fail(`Unknown bundled module without resource: ${module.kind}`);
            }
            if (hash(fs.readFileSync(module.resource.split('?')[0])) !== module.sourceSha256) {
                fail(`Bundled module changed since capture: ${module.resource}`);
            }
            const pkg = packageForResource(module.resource, lock);
            if (pkg) {
                if (pkg.manifestSha256 !== module.manifestSha256) {
                    fail(`Package manifest changed since capture: ${pkg.path}`);
                }
                const current = packages.get(pkg.ref) || { ...pkg, bundles: new Set() };
                chunk.files.forEach(file => current.bundles.add(file));
                packages.set(pkg.ref, current);
                const source = relative(module.resource.split('?')[0]);
                const previous = moduleEvidence.get(source);
                if (previous && (previous.sha256 !== module.sourceSha256 || previous.package !== pkg.ref)) {
                    fail(`Ambiguous module identity: ${source}`);
                }
                const item = previous || { path: source, sha256: module.sourceSha256, package: pkg.ref, bundles: new Set() };
                chunk.files.forEach(file => item.bundles.add(file));
                moduleEvidence.set(source, item);
            }
        }
    }
    const shippedJs = Object.keys(archive.files).filter(name => /^ClientSideAssets\/.*\.js$/.test(name));
    if (shippedJs.length !== files.size || shippedJs.some(name => !files.has(name.slice('ClientSideAssets/'.length)))) {
        fail('SPPKG JavaScript assets do not match the captured chunks.');
    }
    const archiveFiles = {};
    for (const entry of Object.values(archive.files)) {
        if (!entry.dir && !/\.(js|txt|xml|rels|png)$/i.test(entry.name)) {
            fail(`Unclassified SPPKG file: ${entry.name}`);
        }
        if (!entry.dir) {
            if (/\.txt$/i.test(entry.name) &&
                (!entry.name.endsWith('.js.LICENSE.txt') || !archive.file(entry.name.slice(0, -'.LICENSE.txt'.length)))) {
                fail(`Unclassified SPPKG text file: ${entry.name}`);
            }
            archiveFiles[entry.name] = hash(await entry.async('nodebuffer'));
        }
    }
    if (bom.metadata.component.name !== packageJson.name || bom.metadata.component.version !== packageJson.version) {
        fail('Webpack SBOM identity differs from the project.');
    }
    const selected = bom.components.filter(component => packages.has(component['bom-ref']));
    if (selected.length !== packages.size || selected.some(component => !component.purl || !component.licenses?.length)) {
        fail('Bundled package instances are missing from the SBOM or lack purls or licenses.');
    }
    const refs = new Set(selected.map(component => component['bom-ref']));
    if (refs.size !== selected.length || refs.has(bom.metadata.component['bom-ref'])) {
        fail('Duplicate component references in SBOM.');
    }
    const packageSha256 = hash(bytes);
    bom.components = selected;
    bom.dependencies = [
        { ref: bom.metadata.component['bom-ref'], dependsOn: [...refs].sort() },
        ...[...refs].sort().map(ref => ({ ref, dependsOn: [] }))
    ];
    bom.metadata.properties = [
        ...(bom.metadata.properties || []),
        { name: 'spfx-sbom:sppkg:sha256', value: packageSha256 },
        { name: 'spfx-sbom:sppkg:filename', value: path.basename(config.paths.zippedPackage) },
        { name: 'spfx-sbom:evidence:provenance', value: provenance }
    ];
    const validatorPath = require.resolve('@cyclonedx/cyclonedx-library/Validation', {
        paths: [path.dirname(require.resolve('@cyclonedx/webpack-plugin'))]
    });
    const { JsonStrictValidator } = require(validatorPath);
    const serialized = JSON.stringify(bom, null, 2) + '\n';
    const errors = await new JsonStrictValidator('1.6').validate(serialized);
    if (errors !== null) fail(`CycloneDX validation failed: ${JSON.stringify(errors)}`);

    const output = path.join(root, 'SBOM', 'releases', solution.version, packageSha256.slice(0, 12));
    if (fs.existsSync(output)) fail(`SBOM release already exists: ${output}`);
    fs.mkdirSync(output, { recursive: true });
    fs.writeFileSync(path.join(output, 'pnp-script-editor.sppkg.cdx.json'), serialized);
    fs.writeFileSync(path.join(output, path.basename(config.paths.zippedPackage)), bytes);
    fs.writeFileSync(path.join(output, 'evidence.json'), JSON.stringify({
        packageSha256,
        provenance,
        lockSha256: capture.lockSha256,
        bundles: Object.fromEntries(files),
        archiveFiles,
        packages: [...packages.values()].map(pkg => ({ ...pkg, bundles: [...pkg.bundles].sort() })),
        modules: [...moduleEvidence.values()].map(module => ({ ...module, bundles: [...module.bundles].sort() }))
    }, null, 2) + '\n');
    console.log(`Package SBOM: ${output}`);
}

const packagePath = process.argv[2] ? path.resolve(process.argv[2]) :
    path.join(root, 'sharepoint', 'solution', 'pnp-script-editor.sppkg');
const provenance = process.argv[2] ? 'reproduced-bundles-not-original-build' : 'original-build';
generate(packagePath, provenance).catch(error => {
    console.error(error);
    process.exitCode = 1;
});
