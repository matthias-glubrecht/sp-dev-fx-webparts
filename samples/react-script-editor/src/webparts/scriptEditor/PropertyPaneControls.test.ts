jest.mock('react-dom/client', () => ({
    createRoot: jest.fn(() => ({ render: jest.fn(), unmount: jest.fn() }))
}));
jest.mock('@microsoft/sp-core-library', () => ({
    DisplayMode: { Read: 0, Edit: 1 },
    Version: { parse: jest.fn() }
}));
jest.mock('@microsoft/sp-loader', () => ({ SPComponentLoader: { loadScript: jest.fn() } }));
jest.mock('@microsoft/sp-webpart-base', () => ({
    BaseClientSideWebPart: class {
        protected onPropertyPaneFieldChanged(): void {}
        protected onDispose(): void {}
    }
}));
jest.mock('@microsoft/sp-property-pane', () => ({
    PropertyPaneFieldType: { Custom: 1 },
    PropertyPaneLabel: (targetProperty: string, properties: object) => ({ targetProperty, properties }),
    PropertyPaneTextField: (targetProperty: string, properties: object) => ({ targetProperty, properties }),
    PropertyPaneToggle: (targetProperty: string, properties: object) => ({ targetProperty, properties })
}));
jest.mock('@microsoft/sp-http', () => ({ SPHttpClient: { configurations: { v1: {} } } }));
jest.mock('@pnp/sp', () => ({ spfi: jest.fn(), SPFx: jest.fn() }));
jest.mock('@pnp/sp/webs', () => ({}));
jest.mock('@pnp/sp/site-groups/web', () => ({}));
jest.mock('@pnp/queryable', () => ({ Caching: jest.fn() }));
jest.mock('@fluentui/react', () => ({}));
jest.mock('react-ace', () => ({ __esModule: true, default: () => null }));
jest.mock('ace-builds/src-noconflict/mode-html', () => ({}));
jest.mock('ace-builds/src-noconflict/theme-monokai', () => ({}));
jest.mock('./components/ScriptEditor', () => ({ __esModule: true, default: () => null }));

import * as React from 'react';
import { createRoot, type Root } from 'react-dom/client';
import type {
    IPropertyPaneConfiguration,
    IPropertyPaneCustomFieldProps,
    IPropertyPaneField
} from '@microsoft/sp-property-pane';
import type { IAudience } from './IAudience';
import type { IScriptEditorWebPartProps } from './IScriptEditorWebPartProps';
import ScriptEditorWebPart from './ScriptEditorWebPart';

interface IWebPartHarness {
    loadResources(): Promise<void>;
    configuration(): IPropertyPaneConfiguration;
    render(): void;
    dispose(): void;
}

async function createWebPart(): Promise<{ webPart: IWebPartHarness; properties: IScriptEditorWebPartProps }> {
    class TestWebPart extends ScriptEditorWebPart {
        public loadResources(): Promise<void> {
            return this.loadPropertyPaneResources();
        }

        public configuration(): IPropertyPaneConfiguration {
            return this.getPropertyPaneConfiguration();
        }

        public dispose(): void {
            this.onDispose();
        }
    }

    const properties: IScriptEditorWebPartProps = {
        script: 'original',
        useExternalScript: false,
        title: 'Test',
        removePadding: false,
        spPageContextInfo: false,
        teamsContext: false,
        audiences: [],
        audienceCacheDuration: 24
    };
    const webPart = new TestWebPart();
    Object.defineProperties(webPart, {
        properties: { value: properties },
        context: { value: { manifest: { version: '1.0' }, sdks: {}, pageContext: { web: { absoluteUrl: 'https://contoso.sharepoint.com' } } } },
        domElement: { value: document.createElement('div') },
        displayMode: { value: 1, writable: true }
    });
    await webPart.loadResources();
    return { webPart, properties };
}

function getField(webPart: IWebPartHarness, targetProperty: string): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
    const group = webPart.configuration().pages[0].groups[0];
    if (!('groupFields' in group)) {
        throw new Error('Expected a property pane group.');
    }
    const field = group.groupFields.find(item => item.targetProperty === targetProperty);
    if (!field) {
        throw new Error(`Missing property pane field: ${targetProperty}`);
    }
    return field;
}

function lastRoot(createRoot: typeof import('react-dom/client').createRoot): Root {
    const calls = jest.mocked(createRoot).mock.results;
    const result = calls[calls.length - 1];
    if (!result || result.type !== 'return') {
        throw new Error('The custom field did not create a React root.');
    }
    return result.value;
}

function renderedProps<T>(root: Root): T {
    const calls = jest.mocked(root.render).mock.calls;
    const element = calls[calls.length - 1]?.[0];
    if (!React.isValidElement<T>(element)) {
        throw new Error('The custom field did not render a React element.');
    }
    return element.props;
}

describe('property pane fields', () => {
    beforeEach(() => {
        jest.clearAllMocks();
    });

    it('saves script changes through SPFx and reloads the latest value', async () => {
        const { webPart, properties } = await createWebPart();
        const element = document.createElement('div');
        const onChanged = jest.fn((targetProperty?: string, value?: unknown) => {
            if (targetProperty === 'script' && typeof value === 'string') {
                properties.script = value;
            }
        });

        getField(webPart, 'script').properties.onRender(element, undefined, onChanged);
        const root = lastRoot(createRoot);
        const editor = renderedProps<{ initialValue: string; onSave: (value: string) => void }>(root);
        expect(editor.initialValue).toBe('original');
        editor.onSave('<p>updated</p>');
        expect(onChanged).toHaveBeenCalledWith('script', '<p>updated</p>');
        expect(properties.script).toBe('<p>updated</p>');

        const refreshedField = getField(webPart, 'script');
        refreshedField.properties.onRender(element, undefined, onChanged);
        const refreshed = renderedProps<{ initialValue: string; onSave: (value: string) => void }>(root);
        expect(refreshed.initialValue).toBe('<p>updated</p>');
        refreshed.onSave('');
        expect(properties.script).toBe('');
        expect(onChanged).toHaveBeenLastCalledWith('script', '');
        expect(createRoot).toHaveBeenCalledTimes(1);
        refreshedField.properties.onDispose?.(element);
    });

    it('sends audience and cache changes through the property pane callback', async () => {
        const { webPart, properties } = await createWebPart();
        const onChanged = jest.fn((targetProperty?: string, value?: unknown) => {
            if (targetProperty === 'audiences' && Array.isArray(value)) {
                properties.audiences = value;
            } else if (targetProperty === 'audienceCacheDuration' && typeof value === 'number') {
                properties.audienceCacheDuration = value;
            }
        });
        const selected: IAudience[] = [{ id: '7', fullName: 'Members', login: 'Members' }];
        const audienceElement = document.createElement('div');
        const audienceField = getField(webPart, 'audiences');
        audienceField.properties.onRender(audienceElement, undefined, onChanged);
        renderedProps<{ onChange: (value: IAudience[]) => void }>(lastRoot(createRoot)).onChange(selected);
        expect(onChanged).toHaveBeenLastCalledWith('audiences', selected);
        expect(properties.audiences).toEqual(selected);

        const numberElement = document.createElement('div');
        const numberField = getField(webPart, 'audienceCacheDuration');
        numberField.properties.onRender(numberElement, undefined, onChanged);
        renderedProps<{ onChange: (value: number) => void }>(lastRoot(createRoot)).onChange(6);
        expect(onChanged).toHaveBeenLastCalledWith('audienceCacheDuration', 6);
        expect(properties.audienceCacheDuration).toBe(6);

        audienceField.properties.onDispose?.(audienceElement);
        numberField.properties.onDispose?.(numberElement);
    });

    it('reports a missing property pane change callback', async () => {
        const { webPart } = await createWebPart();
        const field = getField(webPart, 'script');
        expect(() => field.properties.onRender(document.createElement('div'))).toThrow(
            'Property pane change callback is missing for script.'
        );
        expect(createRoot).not.toHaveBeenCalled();
    });

    it('reuses an active React root and creates a new one after switching to read mode', async () => {
        const { webPart } = await createWebPart();
        webPart.render();
        await new Promise(resolve => setTimeout(resolve, 0));
        const firstRoot = lastRoot(createRoot);

        webPart.render();
        await new Promise(resolve => setTimeout(resolve, 0));
        expect(createRoot).toHaveBeenCalledTimes(1);
        expect(firstRoot.render).toHaveBeenCalledTimes(2);

        Object.defineProperty(webPart, 'displayMode', { value: 0 });
        webPart.render();
        expect(firstRoot.unmount).toHaveBeenCalledTimes(1);

        Object.defineProperty(webPart, 'displayMode', { value: 1 });
        webPart.render();
        await new Promise(resolve => setTimeout(resolve, 0));
        const secondRoot = lastRoot(createRoot);
        expect(createRoot).toHaveBeenCalledTimes(2);
        expect(secondRoot).not.toBe(firstRoot);
        webPart.dispose();
        expect(secondRoot.unmount).toHaveBeenCalledTimes(1);
    });

    it('restores AMD detection before running inline scripts', async () => {
        const { webPart, properties } = await createWebPart();
        Reflect.deleteProperty(properties, 'audiences');
        const originalDefine = Object.getOwnPropertyDescriptor(window, 'define');
        const originalResult = Object.getOwnPropertyDescriptor(window, '__scriptEditorSawAmd');
        try {
            Object.defineProperty(window, 'define', { value: { amd: {} }, configurable: true, writable: true });
            properties.script = '<script>window.__scriptEditorSawAmd = !!window.define.amd;</script>';
            Object.defineProperty(webPart, 'displayMode', { value: 0 });
            webPart.render();
            expect(Reflect.get(window, '__scriptEditorSawAmd')).toBe(true);
        } finally {
            if (originalDefine) Object.defineProperty(window, 'define', originalDefine);
            else Reflect.deleteProperty(window, 'define');
            if (originalResult) Object.defineProperty(window, '__scriptEditorSawAmd', originalResult);
            else Reflect.deleteProperty(window, '__scriptEditorSawAmd');
        }
    });

    it('does not mount an outdated editor after switching to read mode', async () => {
        const { webPart } = await createWebPart();
        webPart.render();
        Object.defineProperty(webPart, 'displayMode', { value: 0 });
        webPart.render();
        await new Promise(resolve => setTimeout(resolve, 0));
        expect(createRoot).not.toHaveBeenCalled();
    });

    it('does not mount an editor after the web part is disposed', async () => {
        const { webPart } = await createWebPart();
        webPart.render();
        webPart.dispose();
        await new Promise(resolve => setTimeout(resolve, 0));
        expect(createRoot).not.toHaveBeenCalled();
    });

    it('restores AMD detection when an external-script render is superseded', async () => {
        const { webPart, properties } = await createWebPart();
        Reflect.deleteProperty(properties, 'audiences');
        const originalDefine = Object.getOwnPropertyDescriptor(window, 'define');
        const amd = {};
        const define = { amd };
        try {
            Object.defineProperty(window, 'define', { value: define, configurable: true, writable: true });
            properties.script = '<script src="https://example.com/example.js"></script>';
            Object.defineProperty(webPart, 'displayMode', { value: 0 });
            webPart.render();
            expect(define.amd).toBeNull();
            Object.defineProperty(webPart, 'displayMode', { value: 1 });
            webPart.render();
            await new Promise(resolve => setTimeout(resolve, 0));
            expect(define.amd).toBe(amd);
            webPart.dispose();
        } finally {
            if (originalDefine) Object.defineProperty(window, 'define', originalDefine);
            else Reflect.deleteProperty(window, 'define');
        }
    });

    it('keeps an explicitly empty audience hidden in read mode', async () => {
        const { webPart, properties } = await createWebPart();
        properties.script = '<div>Not visible</div>';
        Object.defineProperty(webPart, 'displayMode', { value: 0 });
        webPart.render();
        await new Promise(resolve => setTimeout(resolve, 0));
        const element = Reflect.get(webPart, 'domElement') as HTMLElement;
        expect(element.innerHTML).toBe('');
    });
});
