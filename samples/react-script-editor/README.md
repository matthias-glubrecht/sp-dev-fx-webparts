
# Script editor web part for modern pages

This version works only for SharePoint Online. If you want a version for SharePoint on-premises go to [react-script-editor-onprem](../react-script-editor-onprem).

**It's important you read and understand the notes on [deployment](#deploy-to-non-script-sites--modern-team-sites).**

## Summary

Coming from old classic SharePoint pages you might have existing script solutions you want to re-use on a modern page
without having to repackage it as a new SharePoint Framework web part. This web part is similar to the classic
Script Editor Web Part, and allows you do drop arbitrary script or html on a modern page.

> Notice. All client-side web parts are deployed or enabled to be available in site level by tenant administrator using tenant app catalog. If there are concerns on enabling script options in a tenant, this web part or a approach should not be approved by tenant administrators.

As an example add the following scripts to the web part in order to show stock ticker info on your page. First *tv.js* is loaded, and then the script block is executed to show the ticker information.

```html
<!-- TradingView Widget BEGIN -->
<div class="tradingview-widget-container">
  <div id="tradingview_e7aa0"></div>
  <div class="tradingview-widget-copyright"><a href="https://www.tradingview.com/symbols/NASDAQ-AAPL/" rel="noopener" target="_blank"><span class="blue-text">AAPL Chart</span></a> by TradingView</div>
  <script type="text/javascript" src="https://s3.tradingview.com/tv.js"></script>
  <script type="text/javascript">
  new TradingView.widget(
  {
  "width": 980,
  "height": 610,
  "symbol": "NASDAQ:AAPL",
  "interval": "D",
  "timezone": "Etc/UTC",
  "theme": "light",
  "style": "1",
  "locale": "en",
  "toolbar_bg": "#f1f3f6",
  "enable_publishing": false,
  "allow_symbol_change": true,
  "container_id": "tradingview_e7aa0"
}
  );
  </script>
</div>
<!-- TradingView Widget END -->
```

The web part works by loading each script in a `<script src>` tag sequentially in the order they are specified, then any other `<script>` block is executed.

![Script Editor web part](./assets/modern-script-editor-wp.gif)

If all you want is to add markup on the page, you can do that as well. Adding the following html would show a headline and a list.

```html
<h2>My header</h2>
<ul>
    <li>One
    <li>Two
    <li>Three
</ul>
```

You may add CSS via style tags or `link` tags.

```html
<!-- Latest compiled and minified CSS -->
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" integrity="sha384-BVYiiSIFeK1dGmJRAkycuHAHRg32OmUcww7on3RYdg4Va+PmSTsz/K68vbdEjh4u" crossorigin="anonymous">
<style>
    #container h1 {
        color: red;
    }
</style>
<div id="container">
    <h1>Headline</h1>
</div>
<div class="row">
    <div class="col-md-8">.col-md-8</div>
    <div class="col-md-4">.col-md-4</div>
</div>
<div class="row">
    <div class="col-md-4">.col-md-4</div>
    <div class="col-md-4">.col-md-4</div>
    <div class="col-md-4">.col-md-4</div>
</div>
```

## Support for Office UI Fabric styles

By adding the class name `ms-Fabric` to your top element, you can use use fabric CSS classes directly in the web part. See [Fabric Core](https://developer.microsoft.com/en-us/fabric#/get-started/web#fabric-core) for more information on Fabric classes.

### Sample

```html
<div class="ms-Fabric">
    <span class="ms-font-su ms-fontColor-themePrimary">
        Big text in primary theme color
    </span>
</div>
```

## Support for classic _spPageContextInfo

If your scripts rely on the classic_spPageContextInfo, you can enable that in the web part property pane.

## Support for Teams tabs

If you want to use the solution as a Teams tab, perform the changes to deploy to [non-script](#deploy-to-non-script-sites--modern-team-sites) sites and [tenant wide deployment](#deploy-tenant-wide).

Next see the [Teams tab tutorial](https://learn.microsoft.com/sharepoint/dev/spfx/web-parts/get-started/using-web-part-as-ms-teams-tab) for steps needed to deploy the solution as a Teams tab.

You can get access to the Teams context information via the global variable `_teamsContexInfo`. The available properties are documented at [Teams context interface](https://learn.microsoft.com/javascript/api/@microsoft/teams-js/microsoftteams.context?view=msteams-client-js-latest).

### Sample

```html
<div class="ms-Fabric">
    <span id="ScriptIt" class="ms-font-xxl ms-fontColor-neutralPrimary">
    </span>
</div>
<script>
var element = document.getElementById("ScriptIt");
element.innerHTML = "Team: " + _teamsContexInfo.teamName + "<br\>Channel: " + _teamsContexInfo.channelName;
</script>

```

![Script Editor web part in Teams](./assets/modern-script-editor-wp-teams.gif)

### Deploy to non-script sites / modern team sites

By default this web part is not allowed on no-script sites, as it allows execution of arbitrary script. This is by design as from a security and governance perspective you might not want arbitrary script added to your pages. This is typically something you want control over.

If you however want to allow the web part for non-script sites like Office 365 Group modern team sites you have to edit `ScriptEditorWebPart.manifest.json` with the following change:

```json
"requiresCustomScript": false
```

### Deploy tenant wide

By default you have to install this web part per site collection where you want it available. If you want it enabled by default on all sites you have to edit `package-solution.json` with the following change:

```json
"skipFeatureDeployment": true
```

In order to make it available to absolutely all sites you need apply the *Deploy to non-script sites / modern team site* change as well.

## Property pane controls

The property pane controls are implemented in this sample and do not depend on `@pnp/spfx-property-controls`. The HTML editor uses React Ace; the audience picker searches SharePoint users and groups; and the cache duration field accepts numbers of at least one hour, including fractional hours. Audience selections retain the same property shape used by earlier versions of this web part.

Unlike the previous editor, the replacement does not provide a button to format HTML or scripts.

Because these controls are not the PnP property controls, their PnP-specific telemetry does not apply.

The installed web part version is shown at the top of the property pane.

## Changes in version 1.0.31

The transitive `fast-uri` dependency is pinned to 3.1.9. Ajv 6 is constrained to `>=6.14.0 <7` and resolves to 6.15.0 in the lockfile; Ajv 8 versions are unchanged. No web part source code changed in this release.

### Verified SharePoint package

Download `pnp-script-editor.sppkg`, `pnp-script-editor.sppkg.cdx.json`, and `evidence.json` from [release 1.0.31](https://github.com/matthias-glubrecht/sp-dev-fx-webparts/releases/tag/react-script-editor-v1.0.31). To confirm you have the exact package covered by the SBOM, check its SHA-256 before uploading it to the SharePoint App Catalog:

```powershell
Get-FileHash .\pnp-script-editor.sppkg -Algorithm SHA256
```

The expected SHA-256 is `e4bbea8e389df1680cbee9e4bb2bb075614fc036e8bf1640a5690b20462108e6`. It is also recorded in the SBOM and its evidence file. The separate `pnp-script-editor-1.0.31-dependency-sbom.cdx.json` describes the broader npm lockfile, including build tools; it is not the package-specific SBOM.

### How the package-specific SBOM was generated

`npm run build` uses Heft and Webpack 5 with `@cyclonedx/webpack-plugin` 5.3.3. The generator selects npm package instances from the emitted production chunks, checks the SHA-256 of every JavaScript bundle against the SPPKG, and validates the CycloneDX 1.6 JSON. It saves the SPPKG, SBOM, and evidence (including bundle mappings and archive file hashes) under `SBOM/releases/<solution-version>/<package-sha256-prefix>/`.

For the already-released 1.0.31 package, the JavaScript bundles were rebuilt with Node.js 22.23.1 and matched the four JavaScript files in the release SPPKG byte for byte. The 25 bundled npm package instances also matched the release lockfile by version and integrity. The evidence is explicitly marked as a reproduction, not a capture of the original build. To repeat the check using the SBOM-enabled source on `main`, keep the downloaded release SPPKG in a separate directory so the build does not overwrite it:

```powershell
npm ci
npm run build
npm run test:sbom
node .\scripts\generate-sppkg-sbom.js "C:\Downloads\pnp-script-editor.sppkg"
```

### Snyk scan

The following scans the package-specific SBOM. Run it from the directory containing the downloaded SBOM, or pass its full path. The command sends the SBOM to Snyk and requires a Snyk Enterprise plan.

```powershell
snyk sbom test --file=.\pnp-script-editor.sppkg.cdx.json --json
```

On November 27, 2026, Snyk reported **no known vulnerabilities** across the 24 distinct bundled npm dependencies it identified in the SBOM. This is a point-in-time result for the bundled npm dependencies, not a guarantee that the SPPKG is vulnerability-free. It does not cover the SharePoint host, scripts inserted by page authors, scripts loaded from external URLs, or build-only dependencies.

The separate scans of the development and build dependency tree reported 16 distinct findings. None of the affected npm package versions appears in the package-specific SBOM. Those findings remain relevant to the development and build environment, but do not apply to the dependencies shipped in this SPPKG.

## Changes in version 1.0.30

This version upgrades the previous SPFx 1.23.2 sample rather than creating a new solution. The solution, feature and web part IDs and existing property names remain unchanged so that existing installations and page data can be upgraded.

### React rendering

SPFx and React were updated to 1.24.1 and 18.3.1 respectively. The edit-mode placeholder now uses `createRoot` and `root.render` instead of `ReactDom.render`. A root is reused while the web part remains in edit mode. Before read mode writes script markup directly into the same DOM container, the root is unmounted and discarded; a later return to edit mode creates a new root. Disposal also unmounts the root. This is necessary because an unmounted React 18 root cannot be reused, and React must not retain ownership of a container that is overwritten with `innerHTML`.

Render tokens prevent an older asynchronous audience check or editor import from overwriting a newer render after a mode change. External scripts are still loaded before inline scripts, with AMD detection restored before inline scripts run.

### Property controls and stored data

The previous `@pnp/spfx-property-controls` 3.25.0 targets React 17; this version replaces it with sample-owned controls rather than taking a beta dependency. The HTML editor uses `react-ace` and Ace, the audience picker uses Fluent UI and SharePoint people search, and the cache duration field validates numbers of at least one hour. Their appearance and interaction can differ from the previous controls. The replacement does not offer a Format button: the initial formatter could alter HTML attribute values or significant whitespace, so it was removed to protect script content.

The new custom fields report changes through SPFx's property pane change callback, which updates the property bag and marks the web part as changed. The code editor is configured from the current `script` property whenever the property pane is refreshed. The previous constructor only bound `scriptUpdate`; its property-update role is now handled by the SPFx callback. Existing audience values retain the previous property shape, including user claims, SharePoint group IDs and Entra group claims.

### PnPjs and build

`@pnp/sp` was upgraded from 2.x to 4.23.0. Each web part initializes an SPFx-context-aware `spfi` instance; SharePoint group queries use PnPjs v4's caching behavior. The previous PnP property-control localization resource was removed along with the controls.

The existing Heft toolchain was updated for SPFx 1.24.1, with TypeScript 5.8 and matching React types. `npm run build` now runs the production build and Jest tests before producing the `.sppkg`; the solution and feature versions were raised to 1.0.30.0. The sample-specific dependency overrides are retained in `package.json`.

The web part was tested in a browser on SharePoint Online and the observed result was satisfactory. This was a smoke test, not exhaustive validation of all configurations or existing page data. The automated tests cover audience mapping, property field changes and the React root lifecycle, but use mocked SPFx services and UI components. Additional tenant testing is recommended, particularly for existing page data, people and group selection, mode changes and script execution.

## Compatibility

| :warning: Important          |
|:---------------------------|
| Every SPFx version is optimally compatible with specific versions of Node.js. In order to be able to build this sample, you need to ensure that the version of Node on your workstation matches one of the versions listed in this section. This sample will not work on a different version of Node.|
|Refer to <https://aka.ms/spfx-matrix> for more information on SPFx compatibility.   |

This sample is optimally compatible with the following environment configuration:

![SPFx 1.24.1](https://img.shields.io/badge/SPFx-1.24.1-green.svg)
![Node.js v22, v24, v26](https://img.shields.io/badge/Node.js-v22%20%7C%20v24%20%7C%20v26-green.svg)
![Compatible with SharePoint Online](https://img.shields.io/badge/SharePoint%20Online-Compatible-green.svg)
![Does not work with SharePoint 2019](https://img.shields.io/badge/SharePoint%20Server%202019-Incompatible-red.svg "SharePoint Server 2019 requires SPFx 1.4.1 or lower")
![Does not work with SharePoint 2016 (Feature Pack 2)](https://img.shields.io/badge/SharePoint%20Server%202016%20(Feature%20Pack%202)-Incompatible-red.svg "SharePoint Server 2016 Feature Pack 2 requires SPFx 1.1")
![Local Workbench Unsupported](https://img.shields.io/badge/Local%20Workbench-Unsupported-red.svg "Local workbench is no longer available as of SPFx 1.13 and above")
![Hosted Workbench Deprecated](https://img.shields.io/badge/Hosted%20Workbench-Deprecated-yellow.svg "The hosted workbench is retired on December 1, 2026")
![Compatible with Remote Containers](https://img.shields.io/badge/Remote%20Containers-Compatible-green.svg)

## Applies to

* [SharePoint Framework Release GA](https://blogs.office.com/2017/02/23/sharepoint-framework-reaches-general-availability-build-and-deploy-engaging-web-parts-today/)
* [Microsoft 365 tenant](https://learn.microsoft.com/sharepoint/dev/spfx/set-up-your-development-environment)

## Contributors

* [Mikael Svenson](https://github.com/wobba) 
* [Felix Bohnacker](https://github.com/felixbohnackerfelixbohnacker)
* [salascz](https://github.com/salascz)
* [Waldek Mastykarz](https://github.com/waldekmastykarz)
* [Luccas Castro](https://github.com/devpio)

## Version history

Version|Date|Comments
-------|----|--------
1.0|March 10th, 2017|Initial release
1.0.0.1|August 8th, 2017|Updated SPFx version and CSS loading
1.0.0.2|October 4th, 2017|Updated SPFx version, bundle Office UI Fabric and CSS in web part
1.0.0.3|January 10th, 2018|Updated SPFx version, added remove padding property and refactoring
1.0.0.4|February 14th, 2018|Added title property for edit mode and documentation for enabling the web part on Group sites / tenant wide
1.0.0.5|March 8th, 2018|Added support for loading scripts which are AMD/UMD modules. Added support for classic _spPageContextInfo. Refactoring.
1.0.0.6|March 26th, 2018|Fixed so that AMD modules don't detect `define`, and load as non-modules.
1.0.0.7|May 23rd, 2018|Added supportsFullBleed to manifest.
1.0.0.8|May 23rd, 2018|Updated SPFx to v1.5.1, made editor load dynamically to reduce runtime bundle size, fixed white space issue.
1.0.0.9|Aug 22, 2018|Improved bundle size
1.0.0.10|Jan 16th, 2019|Fix for removing of web part padding
1.0.0.11|March 18th, 2019|Fix for re-loading of script on smart navigation
1.0.0.12|April 15th, 2019|Re-fix for pad removal of web part
1.0.0.13|July 1th, 2019|Downgrade to SPFx v1.4.1 to support SP2019
1.0.0.14|Oct 13th, 2019|Added resolve to fix pnpm issue. Updated author info.
1.0.0.15|Mar 16th, 2020|Upgrade to SPFx v1.10.0. Add support for Teams tab. Renamed package file.
1.0.0.16|April 1st, 2020|Improved how script tags are handled and cleaned up on smart page navigation.
1.0.17.0|January 29th, 2021|Changed versioning to 3 parts. Updated npm packages, restructured documentation, minor change to webpack analyzer setup.
1.0.18.0|June 14, 2022|Upgrade to SPFx v1.13.1
1.0.19.0|August 31, 2022|Added support for section background color
1.0.20.0|October 10, 2022|Added sample html/script with self-executing function
1.0.21.0|March 11, 2023|Bump dependencies to allow react-script-editor to build under SPFx 1.16.1
1.0.22.0|April 24, 2023|Added support for script in external template 
1.0.23.0|September 7, 2023|Upgrade to SPFx 1.17.4
1.0.24.0|March 8, 2024|Upgrade to SPFx 1.18.2
1.0.25.0|September 25, 2024|Upgrade to SPFx 1.19.0, added support for audience targeting using SharePoint Groups, Entra Groups and individuals.
1.0.26.0|October 1, 2024|Upgrade to SPFx 1.20.0
1.0.27.0|November 19, 2025|Upgrade to SPFx 1.21.1 and execute inline scripts using the Function constructor when the CSP permits `unsafe-eval`.
1.0.28.0|November 23, 2026|Upgrade to SPFx 1.22.2 and migrate the build toolchain from Gulp to Heft.
1.0.29.0|November 23, 2026|Upgrade to SPFx 1.23.2.
1.0.30.0|November 23, 2026|Upgrade to SPFx 1.24.1 and React 18, migrate to PnPjs v4, and replace the PnP property controls.
1.0.31.0|November 27, 2026|Update transitive fast-uri and Ajv 6 dependencies.

## Minimal Path to Awesome

### Testing

* Clone this repository
* Use Node.js 22.14.0 or later in the 22.x release line, Node.js 24, or Node.js 26. The `.nvmrc` file selects the version used for this sample's development.
* In the project folder run `npm install`, `npm run build`, and `npx heft start --nobrowser`.
* Open a modern SharePoint test page that you can edit and append `?loadSPFX=true&debugManifestsFile=https://localhost:4321/temp/build/manifests.js` to its URL. If the URL already contains query parameters, use `&` instead of `?`.
* Allow the debug scripts, edit the page, and add the Modern Script Editor web part. See [Debug SharePoint Framework solutions on modern pages](https://learn.microsoft.com/sharepoint/dev/spfx/debug-modern-pages).
* The local HTTPS server requires a trusted development certificate. A dev container exports its CA certificate as `spfx-dev-cert.cer` and `spfx-dev-cert.pem`; trust the appropriate file on the host only if permitted by your organization's policies. See [opening a sample using a development container](https://aka.ms/spfx-devcontainer).

The build also writes a bundle analysis report to `temp/stats/react-script-editor.stats.html`.

### Deploy

* `npm run clean`
* `npx heft build --production`
* `npx heft package-solution --production`
* Upload `.sppkg` file from `sharepoint\solution` to your tenant App Catalog
  * E.g.: https://&lt;tenant&gt;.sharepoint.com/sites/AppCatalog/AppCatalog
* Add the web part to a site collection, and test it on a page

## Features

This web part illustrates the following concepts on top of the SharePoint Framework:

* Re-use existing JavaScript solutions on a modern page
* Office UI Fabric
* React

## Help

We do not support samples, but we this community is always willing to help, and we want to improve these samples. We use GitHub to track issues, which makes it easy for  community members to volunteer their time and help resolve issues.

If you're having issues building the solution, please run [spfx doctor](https://pnp.github.io/cli-microsoft365/cmd/spfx/spfx-doctor/) from within the solution folder to diagnose incompatibility issues with your environment.

You can try looking at [issues related to this sample](https://github.com/pnp/sp-dev-fx-webparts/issues?q=label%3A%22sample%3A%20react-script-editor") to see if anybody else is having the same issues.

You can also try looking at [discussions related to this sample](https://github.com/pnp/sp-dev-fx-webparts/discussions?discussions_q=react-script-editor) and see what the community is saying.

If you encounter any issues while using this sample, [create a new issue](https://github.com/pnp/sp-dev-fx-webparts/issues/new?assignees=&labels=Needs%3A+Triage+%3Amag%3A%2Ctype%3Abug-suspected%2Csample%3A%20react-script-editor&template=bug-report.yml&sample=react-script-editor&authors=@wobba%20@salascz%20@felixbohnacker&title=react-script-editor%20-%20).

For questions regarding this sample, [create a new question](https://github.com/pnp/sp-dev-fx-webparts/issues/new?assignees=&labels=Needs%3A+Triage+%3Amag%3A%2Ctype%3Aquestion%2Csample%3A%20react-script-editor&template=question.yml&sample=react-script-editor&authors=@wobba%20@salascz%20@felixbohnacker&title=react-script-editor%20-%20).

Finally, if you have an idea for improvement, [make a suggestion](https://github.com/pnp/sp-dev-fx-webparts/issues/new?assignees=&labels=Needs%3A+Triage+%3Amag%3A%2Ctype%3Aenhancement%2Csample%3A%20react-script-editor&template=question.yml&sample=react-script-editor&authors=@wobba%20@salascz%20@felixbohnacker&title=react-script-editor%20-%20).

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

<img src="https://m365-visitor-stats.azurewebsites.net/sp-dev-fx-webparts/samples/react-script-editor" />
