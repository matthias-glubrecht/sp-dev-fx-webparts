'use strict';

const path = require('path');
const { BundleAnalyzerPlugin } = require('webpack-bundle-analyzer');

module.exports = function (webpackConfig) {
    const projectPath = path.resolve(__dirname, '../..');
    const dropPath = path.join(projectPath, 'temp', 'stats');

    webpackConfig.plugins.push(new BundleAnalyzerPlugin({
        openAnalyzer: false,
        analyzerMode: 'static',
        reportFilename: path.join(dropPath, `${path.basename(projectPath)}.stats.html`),
        generateStatsFile: false,
        logLevel: 'error'
    }));

    return webpackConfig;
};
