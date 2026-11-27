'use strict';

const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawnSync } = require('node:child_process');
const test = require('node:test');
const JSZip = require('jszip');

const root = path.resolve(__dirname, '..');
const packagePath = path.join(root, 'sharepoint', 'solution', 'pnp-script-editor.sppkg');
const packageBytes = fs.readFileSync(packagePath);
const digest = crypto.createHash('sha256').update(packageBytes).digest('hex');
const version = JSON.parse(fs.readFileSync(path.join(root, 'config', 'package-solution.json'))).solution.version;
const output = path.join(root, 'SBOM', 'releases', version, digest.slice(0, 12));

test('inventories lazy-loaded packages without build-only or unused packages', () => {
    const bom = JSON.parse(fs.readFileSync(path.join(output, 'pnp-script-editor.sppkg.cdx.json')));
    const names = new Set(bom.components.map(component => `${component.group ? `${component.group}/` : ''}${component.name}`));
    assert(names.has('ace-builds'));
    assert(names.has('@pnp/sp'));
    assert(!names.has('ajv'));
    assert(!names.has('proxy-addr'));
    assert(!names.has('@fluentui/date-time-utilities'));
});

for (const scenario of ['modified bundle', 'unknown bundle']) {
    test(`rejects ${scenario}`, async t => {
        const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'spfx-sbom-test-'));
        const fixture = path.join(directory, 'fixture.sppkg');
        t.after(() => {
            if (fs.existsSync(fixture)) fs.unlinkSync(fixture);
            fs.rmdirSync(directory);
        });
        const zip = await JSZip.loadAsync(packageBytes);
        if (scenario === 'modified bundle') {
            const name = Object.keys(zip.files).find(file => /^ClientSideAssets\/.*\.js$/.test(file));
            assert(name);
            zip.file(name, 'modified');
        } else {
            zip.file('ClientSideAssets/unexpected.js', 'unexpected');
        }
        fs.writeFileSync(fixture, await zip.generateAsync({ type: 'nodebuffer' }));
        const result = spawnSync(process.execPath, [path.join(root, 'scripts', 'generate-sppkg-sbom.js'), fixture], {
            cwd: root,
            encoding: 'utf8'
        });
        assert.notEqual(result.status, 0);
        assert.match(result.stderr, /SPPKG bundle does not match|JavaScript assets do not match/);
    });
}
