import * as React from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IPropertyPaneConfiguration, IPropertyPaneField, PropertyPaneLabel, PropertyPaneTextField, PropertyPaneToggle } from "@microsoft/sp-property-pane";
import { spfi, SPFx, type SPFI } from '@pnp/sp';
import { IScriptEditorProps } from './components/IScriptEditorProps';
import { IScriptEditorWebPartProps } from './IScriptEditorWebPartProps';
import PropertyPaneLogo from './PropertyPaneLogo';
import { UserGroupCheck } from './AccessCheck';
import '@pnp/sp/webs';
import '@pnp/sp/site-groups/web';

export default class ScriptEditorWebPart extends BaseClientSideWebPart<IScriptEditorWebPartProps> {
    private _propertyPaneControls: typeof import('./PropertyPaneControls') | undefined;
    private _externalScriptContent: string | undefined;
    private _root: Root | undefined;
    private _renderToken = 0;
    private _sp: SPFI;

    protected async onInit(): Promise<void> {
        if (this.properties.useExternalScript) {
            try {
                const prefix = this.properties.externalScript.indexOf('?') === -1 ? '?' : '&';
                const response = await fetch(`${this.properties.externalScript}${prefix}pnp=${new Date().getTime()}`);
                this._externalScriptContent = await response.text();
            } catch {
                this._externalScriptContent = 'Failed to load external script.';
            }
        }

        await super.onInit();
        this._sp = spfi().using(SPFx(this.context));
    }

    public render(): void {
        const token = ++this._renderToken;
        void this.renderAsync(token).catch((error) => {
            console.error('Failed to render the Script Editor:', error);
        });
    }

    private async renderAsync(token: number): Promise<void> {
        if (this.displayMode === DisplayMode.Read) {
            if (this.properties.removePadding) {
                let element = this.domElement.parentElement;
                // check up to 5 levels up for padding and exit once found
                for (let i = 0; i < 5 && element; i++) {
                    const style = window.getComputedStyle(element);
                    const hasPadding = style.paddingTop !== "0px";
                    if (hasPadding) {
                        element.style.paddingTop = "0px";
                        element.style.paddingBottom = "0px";
                        element.style.marginTop = "0px";
                        element.style.marginBottom = "0px";
                    }
                    element = element.parentElement;
                }
            }

            this.unmountReact();
            this.domElement.innerHTML = '';
            if (this.properties.audiences) {
                const checker = new UserGroupCheck(this.properties.audiences, this.properties.audienceCacheDuration, this.context, this._sp);
                const isInAudience = await checker.CheckAudiences();
                if (token !== this._renderToken) return;
                if (!isInAudience) return;
            }

            if (token !== this._renderToken) return;
            this.domElement.innerHTML = this.properties.useExternalScript
                ? (this._externalScriptContent || '')
                : (this.properties.script || '');
            await this.executeScript(this.domElement, token);
        } else {
            await this.renderEditor(token);
        }
    }

    private async renderEditor(token: number): Promise<void> {
        // Dynamically load the editor pane to reduce overall bundle size
        const editorPopUp = await import(
            /* webpackChunkName: 'scripteditor' */
            './components/ScriptEditor'
        );
        if (token !== this._renderToken || this.displayMode === DisplayMode.Read) return;

        if (!this._root) {
            this.domElement.innerHTML = '';
            this._root = createRoot(this.domElement);
        }
        const element: React.ReactElement<IScriptEditorProps> = React.createElement(
            editorPopUp.default,
            {
                script: this.properties.script,
                title: this.properties.title,
                propPaneHandle: this.context.propertyPane,
                key: "pnp" + new Date().getTime()
            }
        );
        this._root.render(element);
    }

    private unmountReact(): void {
        if (this._root) {
            this._root.unmount();
            this._root = undefined;
        }
    }

    protected onDispose(): void {
        this._renderToken++;
        this.unmountReact();
        super.onDispose();
    }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    protected async loadPropertyPaneResources(): Promise<void> {
        this._propertyPaneControls = await import(
            /* webpackChunkName: 'propertypane-fields' */
            './PropertyPaneControls'
        );
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        const controls = this.propertyPaneControls;
        const webPartOptions: IPropertyPaneField<any>[] = [
            PropertyPaneLabel('webPartVersion', {
                text: `Web part version: ${this.context.manifest.version}`
            }),
            PropertyPaneTextField("title", {
                label: "Title to show in edit mode",
                value: this.properties.title
            }),
            PropertyPaneToggle("removePadding", {
                label: "Remove top/bottom padding of web part container",
                checked: this.properties.removePadding,
                onText: "Remove padding",
                offText: "Keep padding"
            }),
            PropertyPaneToggle("spPageContextInfo", {
                label: "Enable classic _spPageContextInfo",
                checked: this.properties.spPageContextInfo,
                onText: "Enabled",
                offText: "Disabled"
            }),
            PropertyPaneToggle('useExternalScript', {
                label: 'Use script from an external URL instead of inline script',
                onText: "Use external script",
                offText: "Use inline script"
            }),
            controls.PropertyPaneAudiencePicker(
                'audiences',
                this.context,
                this.properties.audiences || []
            ),
            controls.PropertyPaneNumberField(
                'audienceCacheDuration',
                'Audience cache duration in hours',
                'Duration in hours to cache the audience information',
                this.properties.audienceCacheDuration || 24
            )
        ];

        if (this.properties.useExternalScript) {
            webPartOptions.push(PropertyPaneTextField("externalScript", {
                label: "Script URL",
                value: this.properties.externalScript,
            }));
        } else {
            webPartOptions.push(controls.PropertyPaneCodeEditor('script', {
                label: 'Edit HTML Code',
                panelTitle: 'Edit HTML Code',
                initialValue: this.properties.script
            }));
        }

        if (this.context.sdks.microsoftTeams) {
            const config = PropertyPaneToggle("teamsContext", {
                label: "Enable teams context as _teamsContexInfo",
                checked: this.properties.teamsContext,
                onText: "Enabled",
                offText: "Disabled"
            });
            webPartOptions.push(config);
        }
        webPartOptions.push(new PropertyPaneLogo());

        return {
            pages: [
                {
                    groups: [
                        {
                            groupFields: webPartOptions
                        }
                    ]
                }
            ]
        };
    }

    private get propertyPaneControls(): typeof import('./PropertyPaneControls') {
        if (!this._propertyPaneControls) {
            throw new Error('Property pane controls have not been loaded.');
        }
        return this._propertyPaneControls;
    }


    private evalScript(elem): void {
        const data = (elem.text || elem.textContent || elem.innerHTML || "");
        
        // Check if this is an external script
        if (elem.src && elem.src.length > 0) {
            // External scripts are handled in executeScript method
            return;
        }

        // For inline scripts, use Function constructor which works with CSP 'unsafe-eval'
        if (data && data.trim().length > 0) {
            try {
                // Use Function constructor instead of eval for better scoping
                // This works because SharePoint CSP includes 'unsafe-eval'
                // eslint-disable-next-line no-new-func -- Inline scripts require CSP to allow unsafe-eval.
                const scriptFunc = new Function(data);
                scriptFunc.call(window);
            } catch (error) {
                if (console.error) {
                    console.error('Failed to execute inline script:', error);
                }
            }
        }
    }

    // Finds and executes scripts in a newly added element's body.
    // Needed since innerHTML does not run scripts.
    //
    // Argument element is an element in the dom.
    private async executeScript(element: HTMLElement, token: number): Promise<void> {

        if (this.properties.spPageContextInfo && !window["_spPageContextInfo"]) {
            window["_spPageContextInfo"] = this.context.pageContext.legacyPageContext;
        }

        if (this.properties.teamsContext && !window["_teamsContexInfo"]) {
            window["_teamsContexInfo"] = this.context.sdks.microsoftTeams.context;
        }

        // Define global name to tack scripts on in case script to be loaded is not AMD/UMD
        (<any>window).ScriptGlobal = {};

        // main section of function
        const scripts = [];
        const children_nodes = element.getElementsByTagName("script");

        for (let i = 0; children_nodes[i]; i++) {
            const child: any = children_nodes[i];
            if (!child.type || child.type.toLowerCase() === "text/javascript") {
                scripts.push(child);
            }
        }

        const urls = [];
        const onLoads = [];
        for (let i = 0; scripts[i]; i++) {
            const scriptTag = scripts[i];
            if (scriptTag.src && scriptTag.src.length > 0) {
                urls.push(scriptTag.src);
            }
            if (scriptTag.onload && scriptTag.onload.length > 0) {
                onLoads.push(scriptTag.onload);
            }
        }

        const amdWindow = window as Window & { define?: { amd?: unknown } };
        let oldamd: unknown = null;
        if (amdWindow.define && amdWindow.define.amd) {
            oldamd = amdWindow.define.amd;
            amdWindow.define.amd = null;
        }

        try {
            for (let i = 0; i < urls.length; i++) {
                try {
                    let scriptUrl = urls[i];
                    // Add unique param to force load on each run to overcome smart navigation in the browser as needed
                    const prefix = scriptUrl.indexOf('?') === -1 ? '?' : '&';
                    scriptUrl += prefix + 'pnp=' + new Date().getTime();
                    await SPComponentLoader.loadScript(scriptUrl, { globalExportsName: "ScriptGlobal" });
                } catch (error) {
                    console.error(error);
                }
                if (token !== this._renderToken) return;
            }

            if (oldamd) {
                amdWindow.define.amd = oldamd;
                oldamd = null;
            }

            for (let i = 0; scripts[i]; i++) {
                const scriptTag = scripts[i];
                if (scriptTag.parentNode) { scriptTag.parentNode.removeChild(scriptTag); }
                this.evalScript(scripts[i]);
            }
            // execute any onload people have added
            for (let i = 0; onLoads[i]; i++) {
                onLoads[i]();
            }
        } finally {
            if (oldamd) {
                amdWindow.define.amd = oldamd;
            }
        }
    }
}
