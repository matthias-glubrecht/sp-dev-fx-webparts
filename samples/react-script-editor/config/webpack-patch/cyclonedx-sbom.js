'use strict';

const { CycloneDxWebpackPlugin } = require('@cyclonedx/webpack-plugin');
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');

module.exports = function (webpackConfig) {
    if (webpackConfig.mode === 'production') {
        webpackConfig.plugins.push(new CycloneDxWebpackPlugin({
            specVersion: '1.6',
            reproducibleResults: true,
            includeWellknown: false,
            outputLocation: 'sbom'
        }));
        webpackConfig.plugins.push({
            apply(compiler) {
                compiler.hooks.thisCompilation.tap('CaptureBundledModules', compilation => {
                    compilation.hooks.processAssets.tap({
                        name: 'CaptureBundledModules',
                        stage: compiler.webpack.Compilation.PROCESS_ASSETS_STAGE_REPORT
                    }, () => {
                        const chunks = [];
                        for (const chunk of compilation.chunks) {
                            const modules = [];
                            const collect = module => {
                                const nested = module.modules ? Array.from(module.modules) : [];
                                const resource = module.resource;
                                let sourceSha256;
                                let manifestSha256;
                                if (resource) {
                                    const file = resource.split('?')[0];
                                    sourceSha256 = crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex');
                                    if (file.includes(`${path.sep}node_modules${path.sep}`)) {
                                        let dir = path.dirname(file);
                                        while (dir !== path.dirname(dir)) {
                                            const manifest = path.join(dir, 'package.json');
                                            if (fs.existsSync(manifest)) {
                                                manifestSha256 = crypto.createHash('sha256').update(fs.readFileSync(manifest)).digest('hex');
                                                break;
                                            }
                                            dir = path.dirname(dir);
                                        }
                                        if (!manifestSha256) throw new Error(`Missing package manifest: ${file}`);
                                    }
                                }
                                modules.push({
                                    type: module.type,
                                    kind: module.constructor.name,
                                    resource,
                                    sourceSha256,
                                    manifestSha256,
                                    nestedCount: nested.length
                                });
                                for (const child of nested) collect(child);
                            };
                            for (const module of compilation.chunkGraph.getChunkModulesIterable(chunk)) collect(module);
                            const files = Array.from(chunk.files).filter(file => file.endsWith('.js'));
                            const hashes = {};
                            for (const file of files) {
                                const asset = compilation.getAsset(file);
                                if (!asset) throw new Error(`Missing webpack asset: ${file}`);
                                hashes[file] = crypto.createHash('sha256').update(asset.source.source()).digest('hex');
                            }
                            chunks.push({ files, hashes, modules });
                        }
                        const output = path.join(compiler.context, 'temp', 'sbom');
                        fs.mkdirSync(output, { recursive: true });
                        const bomName = path.join('sbom', 'bom.json');
                        const bom = compilation.getAsset(bomName);
                        if (!bom) throw new Error('Webpack SBOM was not generated.');
                        fs.writeFileSync(path.join(output, 'raw-bom.json'), bom.source.source());
                        for (const name of [bomName, path.join('sbom', 'bom.xml')]) {
                            if (compilation.getAsset(name)) compilation.deleteAsset(name);
                        }
                        const lock = fs.readFileSync(path.join(compiler.context, 'package-lock.json'));
                        fs.writeFileSync(path.join(output, 'modules.json'), JSON.stringify({
                            lockSha256: crypto.createHash('sha256').update(lock).digest('hex'),
                            chunks
                        }));
                    });
                });
            }
        });
    }

    return webpackConfig;
};
