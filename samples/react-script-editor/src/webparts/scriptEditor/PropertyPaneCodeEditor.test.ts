jest.mock('@microsoft/sp-property-pane', () => ({ PropertyPaneFieldType: { Custom: 1 } }));
jest.mock('@microsoft/sp-http', () => ({ SPHttpClient: { configurations: { v1: {} } } }));
jest.mock('@fluentui/react', () => {
    const React: typeof import('react') = jest.requireActual('react');
    const Button = (props: { text?: string; ariaLabel?: string; onClick?: () => void }): React.ReactElement =>
        React.createElement('button', { 'aria-label': props.ariaLabel, onClick: props.onClick }, props.text || props.ariaLabel);
    return {
        Label: (props: { children?: React.ReactNode }) => React.createElement('label', null, props.children),
        TextField: (props: { value?: string; ariaLabel?: string; onClick?: () => void; type?: string;
            onChange?: (event: React.FormEvent<HTMLInputElement>, value: string) => void }) =>
            React.createElement('input', {
                value: props.value, 'aria-label': props.ariaLabel, onClick: props.onClick, type: props.type,
                readOnly: !props.onChange,
                onChange: (event: React.ChangeEvent<HTMLInputElement>) => props.onChange?.(event, event.currentTarget.value)
            }),
        IconButton: Button,
        PrimaryButton: Button,
        DefaultButton: Button,
        MessageBar: (props: { children?: React.ReactNode }) => React.createElement('div', null, props.children),
        MessageBarType: { error: 1 },
        NormalPeoplePicker: (props: { onResolveSuggestions: (query: string) => Promise<{ text?: string }[]> }) => {
            const [result, setResult] = React.useState('Search');
            return React.createElement('button', {
                onClick: () => { void props.onResolveSuggestions('abc').then(items => setResult(items[0]?.text || 'No results')); }
            }, result);
        },
        PanelType: { large: 1 },
        Panel: (props: { isOpen: boolean; children?: React.ReactNode; onRenderFooterContent: () => React.ReactNode }) =>
            props.isOpen ? React.createElement('div', null, props.children, props.onRenderFooterContent()) : null
    };
});
jest.mock('react-ace', () => {
    const React: typeof import('react') = jest.requireActual('react');
    return {
        __esModule: true,
        default: (props: { value: string; onChange: (value: string) => void }) =>
            React.createElement('textarea', {
                value: props.value,
                onChange: (event: React.ChangeEvent<HTMLTextAreaElement>) => props.onChange(event.currentTarget.value)
            })
    };
});
jest.mock('ace-builds/src-noconflict/mode-html', () => ({}));
jest.mock('ace-builds/src-noconflict/theme-monokai', () => ({}));

import * as React from 'react';
import { flushSync } from 'react-dom';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import type { IAudience } from './IAudience';
import { PropertyPaneAudiencePicker, PropertyPaneCodeEditor, PropertyPaneNumberField } from './PropertyPaneControls';

function button(container: HTMLElement, label: string): HTMLButtonElement {
    const result = Array.from(container.querySelectorAll('button')).find(
        element => element.textContent === label || element.getAttribute('aria-label') === label
    );
    if (!result) {
        throw new Error(`Missing button: ${label}`);
    }
    return result;
}

function editor(container: HTMLElement): HTMLTextAreaElement {
    const result = container.querySelector('textarea');
    if (!result) {
        throw new Error('The code editor is not open.');
    }
    return result;
}

function enterCode(element: HTMLTextAreaElement, value: string): void {
    const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')?.set;
    if (!setter) {
        throw new Error('Textarea value setter is unavailable.');
    }
    flushSync(() => {
        setter.call(element, value);
        element.dispatchEvent(new Event('input', { bubbles: true }));
    });
}

describe('code editor property field', () => {
    it('saves edits, reopens with the saved value, and discards cancelled edits', () => {
        const container = document.createElement('div');
        document.body.appendChild(container);
        let script = 'original';
        const onChanged = jest.fn((targetProperty?: string, value?: unknown) => {
            if (targetProperty === 'script' && typeof value === 'string') {
                script = value;
            }
        });
        const field = (): ReturnType<typeof PropertyPaneCodeEditor> => PropertyPaneCodeEditor('script', {
            label: 'Edit HTML Code',
            panelTitle: 'Edit HTML Code',
            initialValue: script
        });
        let currentField = field();
        flushSync(() => currentField.properties.onRender(container, undefined, onChanged));
        flushSync(() => button(container, 'Edit HTML').click());
        expect(editor(container).value).toBe('original');

        enterCode(editor(container), '<p>updated</p>');
        expect(onChanged).not.toHaveBeenCalled();
        flushSync(() => button(container, 'Save').click());
        expect(onChanged).toHaveBeenCalledWith('script', '<p>updated</p>');
        expect(script).toBe('<p>updated</p>');

        currentField = field();
        flushSync(() => currentField.properties.onRender(container, undefined, onChanged));
        flushSync(() => button(container, 'Edit HTML').click());
        expect(editor(container).value).toBe('<p>updated</p>');
        enterCode(editor(container), 'discarded');
        flushSync(() => button(container, 'Cancel').click());
        flushSync(() => button(container, 'Edit HTML').click());
        expect(editor(container).value).toBe('<p>updated</p>');
        expect(onChanged).toHaveBeenCalledTimes(1);

        enterCode(editor(container), '');
        flushSync(() => button(container, 'Save').click());
        expect(onChanged).toHaveBeenLastCalledWith('script', '');
        currentField = field();
        flushSync(() => currentField.properties.onRender(container, undefined, onChanged));
        flushSync(() => button(container, 'Edit HTML').click());
        expect(editor(container).value).toBe('');

        flushSync(() => currentField.properties.onDispose(container));
        container.remove();
    });

    it('accepts fractional cache hours but rejects values below one', () => {
        const container = document.createElement('div');
        document.body.appendChild(container);
        const onChanged = jest.fn();
        const field = PropertyPaneNumberField('audienceCacheDuration', 'Cache', 'Hours', 24);
        flushSync(() => field.properties.onRender(container, undefined, onChanged));
        const input = container.querySelector('input');
        const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
        if (!input || !setter) throw new Error('Number field is unavailable.');
        flushSync(() => {
            setter.call(input, '0.5');
            input.dispatchEvent(new Event('input', { bubbles: true }));
        });
        expect(onChanged).not.toHaveBeenCalled();
        flushSync(() => {
            setter.call(input, '1.5');
            input.dispatchEvent(new Event('input', { bubbles: true }));
        });
        expect(onChanged).toHaveBeenCalledWith('audienceCacheDuration', 1.5);
        flushSync(() => field.properties.onDispose(container));
        container.remove();
    });

    it('searches with a legacy selected group that has no login', async () => {
        const container = document.createElement('div');
        document.body.appendChild(container);
        const context: WebPartContext = Object.create(null);
        Object.defineProperties(context, {
            pageContext: { value: { web: { absoluteUrl: 'https://contoso.sharepoint.com' } } },
            spHttpClient: { value: {
                post: jest.fn().mockResolvedValue({
                    ok: true,
                    json: async () => ({
                        value: JSON.stringify([{ Key: 'user', DisplayText: 'Jane', Description: 'jane@example.com', EntityType: 'User' }])
                    })
                })
            } }
        });
        const legacySelection = JSON.parse('[{"id":"7","fullName":"Members"}]') as IAudience[];
        const field = PropertyPaneAudiencePicker('audiences', context, legacySelection);
        flushSync(() => field.properties.onRender(container, undefined, jest.fn()));
        flushSync(() => button(container, 'Search').click());
        await new Promise(resolve => setTimeout(resolve, 0));
        expect(button(container, 'Jane')).toBeTruthy();
        flushSync(() => field.properties.onDispose(container));
        container.remove();
    });
});
