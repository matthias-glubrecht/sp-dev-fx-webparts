import * as React from 'react';
import { createRoot, type Root } from 'react-dom/client';
import {
    DefaultButton,
    IconButton,
    Label,
    MessageBar,
    MessageBarType,
    NormalPeoplePicker,
    Panel,
    PanelType,
    PrimaryButton,
    TextField,
    type IPersonaProps
} from '@fluentui/react';
import type { IPropertyPaneCustomFieldProps, IPropertyPaneField } from '@microsoft/sp-property-pane';
import { PropertyPaneFieldType } from '@microsoft/sp-property-pane';
import { SPHttpClient } from '@microsoft/sp-http';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import AceEditor from 'react-ace';
import 'ace-builds/src-noconflict/mode-html';
import 'ace-builds/src-noconflict/theme-monokai';
import { toAudience, type IAudience, type ISharePointPickerEntity } from './IAudience';

interface ICustomFieldProperties extends IPropertyPaneCustomFieldProps {
    onDispose: (element: HTMLElement) => void;
}

const propertyPaneRoots = new WeakMap<HTMLElement, Root>();

interface ICodeEditorProps {
    label: string;
    panelTitle: string;
    initialValue: string;
    onSave: (value: string) => void;
}

interface ICodeEditorState {
    code: string;
    isOpen: boolean;
}

class CodeEditor extends React.Component<ICodeEditorProps, ICodeEditorState> {
    public constructor(props: ICodeEditorProps) {
        super(props);
        this.state = { code: props.initialValue || '', isOpen: false };
    }

    public componentDidUpdate(previousProps: ICodeEditorProps): void {
        if (previousProps.initialValue !== this.props.initialValue && !this.state.isOpen) {
            this.setState({ code: this.props.initialValue || '' });
        }
    }

    private readonly open = (): void => {
        this.setState({ code: this.props.initialValue || '', isOpen: true });
    };

    private readonly close = (): void => {
        this.setState({ code: this.props.initialValue || '', isOpen: false });
    };

    private readonly save = (): void => {
        this.props.onSave(this.state.code);
        this.setState({ isOpen: false });
    };

    public render(): React.ReactElement {
        return (
            <>
                <Label>{this.props.label}</Label>
                <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                    <TextField
                        readOnly
                        value={this.props.initialValue}
                        onClick={this.open}
                        ariaLabel={this.props.label}
                        styles={{ root: { flexGrow: 1 } }}
                    />
                    <IconButton
                        iconProps={{ iconName: 'Code' }}
                        onClick={this.open}
                        ariaLabel="Edit HTML"
                    />
                </div>
                <Panel
                    isOpen={this.state.isOpen}
                    type={PanelType.large}
                    headerText={this.props.panelTitle}
                    isLightDismiss
                    onDismiss={this.close}
                    onRenderFooterContent={() => (
                        <div>
                            <PrimaryButton text="Save" onClick={this.save} />
                            <DefaultButton text="Cancel" onClick={this.close} />
                        </div>
                    )}
                    isFooterAtBottom
                >
                    <AceEditor
                        mode="html"
                        theme="monokai"
                        name="script-editor-html"
                        value={this.state.code}
                        onChange={(code: string) => this.setState({ code })}
                        editorProps={{ $blockScrolling: true }}
                        width="100%"
                        height="70vh"
                        showPrintMargin={false}
                        setOptions={{ useWorker: false, wrap: true }}
                    />
                </Panel>
            </>
        );
    }
}

interface INumberFieldProps {
    label: string;
    description: string;
    value: number;
    onChange: (value: number) => void;
}

interface INumberFieldState {
    value: string;
    error?: string;
}

class NumberField extends React.Component<INumberFieldProps, INumberFieldState> {
    public constructor(props: INumberFieldProps) {
        super(props);
        this.state = { value: String(props.value) };
    }

    public componentDidUpdate(previousProps: INumberFieldProps): void {
        if (previousProps.value !== this.props.value) {
            this.setState({ value: String(this.props.value), error: undefined });
        }
    }

    private readonly onChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
        const text = value || '';
        const parsed = Number(text);
        if (text.trim() !== '' && Number.isFinite(parsed) && parsed >= 1) {
            this.setState({ value: text, error: undefined });
            this.props.onChange(parsed);
        } else {
            this.setState({
                value: text,
                error: 'Enter a number greater than or equal to 1.'
            });
        }
    };

    public render(): React.ReactElement {
        return (
            <TextField
                label={this.props.label}
                description={this.props.description}
                type="number"
                min={1}
                step="any"
                value={this.state.value}
                onChange={this.onChange}
                errorMessage={this.state.error}
            />
        );
    }
}

interface ISharePointPickerResponse {
    value: string;
}

interface IPeoplePickerProps {
    context: WebPartContext;
    label: string;
    selected: IAudience[];
    onChange: (value: IAudience[]) => void;
}

interface IPeoplePickerState {
    selected: IAudience[];
    searchError?: string;
}

const principalTypes = 13;

function initials(name: string): string {
    return name.split(/\s+/).slice(0, 2).map(part => part.charAt(0)).join('');
}

function toPersona(audience: IAudience): IPersonaProps {
    return {
        key: audience.id || audience.login,
        text: audience.fullName,
        secondaryText: audience.jobTitle,
        imageUrl: audience.imageUrl,
        imageInitials: audience.initials || initials(audience.fullName)
    };
}

class AudiencePicker extends React.Component<IPeoplePickerProps, IPeoplePickerState> {
    private _searchSequence = 0;
    private _isDisposed = false;

    public constructor(props: IPeoplePickerProps) {
        super(props);
        this.state = { selected: props.selected || [] };
    }

    public componentDidUpdate(previousProps: IPeoplePickerProps): void {
        if (previousProps.selected !== this.props.selected && !this._isDisposed) {
            this.setState({ selected: this.props.selected || [] });
        }
    }

    public componentWillUnmount(): void {
        this._isDisposed = true;
        this._searchSequence++;
    }

    private readonly search = async (query: string, sequence: number): Promise<IAudience[]> => {
        if (query.trim().length < 3) {
            return [];
        }

        const url = `${this.props.context.pageContext.web.absoluteUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;
        const response = await this.props.context.spHttpClient.post(
            url,
            SPHttpClient.configurations.v1,
            {
                headers: {
                    accept: 'application/json',
                    'content-type': 'application/json'
                },
                body: JSON.stringify({
                    queryParams: {
                        AllowEmailAddresses: true,
                        AllowMultipleEntities: false,
                        AllUrlZones: false,
                        MaximumEntitySuggestions: 20,
                        PrincipalSource: 15,
                        PrincipalType: principalTypes,
                        QueryString: query
                    }
                })
            }
        );
        if (!response.ok) {
            throw new Error(`People search failed with HTTP ${response.status}.`);
        }

        const result = await response.json() as ISharePointPickerResponse;
        let entities: ISharePointPickerEntity[];
        try {
            entities = JSON.parse(result.value) as ISharePointPickerEntity[];
        } catch (error) {
            const detail = error instanceof Error ? error.message : String(error);
            throw new Error(`SharePoint returned an invalid people search response: ${detail}`);
        }

        const selectedKeys = new Set(this.state.selected.map(person =>
            `${(person.login || '').toLocaleLowerCase()}|${person.id || ''}`
        ));
        const audiences = entities
            .filter(entity => typeof entity.Key === 'string' && entity.EntityData?.PrincipalType !== 'UNVALIDATED_EMAIL_ADDRESS')
            .map(entity => toAudience(entity, this.props.context))
            .filter((person): person is IAudience => !!person)
            .filter(person => !selectedKeys.has(`${(person.login || '').toLocaleLowerCase()}|${person.id || ''}`));

        if (sequence !== this._searchSequence || this._isDisposed) {
            return [];
        }
        this.setState({ searchError: undefined });
        return audiences;
    };

    private readonly onChange = (items?: IPersonaProps[]): void => {
        const personas = items || [];
        const selectedByKey = new Map(
            [...this.state.selected, ...this._lastSearchResults].map(person => [person.id || person.login, person])
        );
        const selected = personas
            .map(persona => selectedByKey.get(String(persona.key)))
            .filter((person): person is IAudience => !!person);
        this.setState({ selected, searchError: undefined });
        this.props.onChange(selected);
    };

    private _lastSearchResults: IAudience[] = [];

    private readonly resolveSuggestions = async (query: string): Promise<IPersonaProps[]> => {
        const sequence = ++this._searchSequence;
        try {
            this._lastSearchResults = await this.search(query, sequence);
            return this._lastSearchResults.map(toPersona);
        } catch (error) {
            if (sequence === this._searchSequence && !this._isDisposed) {
                this.setState({ searchError: error instanceof Error ? error.message : String(error) });
            }
            return [];
        }
    };

    public render(): React.ReactElement {
        return (
            <div>
                <Label>{this.props.label}</Label>
                <NormalPeoplePicker
                    onResolveSuggestions={this.resolveSuggestions}
                    onChange={this.onChange}
                    selectedItems={this.state.selected.map(toPersona)}
                    pickerSuggestionsProps={{
                        suggestionsHeaderText: 'Suggested people and groups',
                        noResultsFoundText: 'No people or groups found',
                        loadingText: 'Searching…'
                    }}
                    itemLimit={undefined}
                />
                {this.state.searchError && (
                    <MessageBar messageBarType={MessageBarType.error}>
                        {this.state.searchError}
                    </MessageBar>
                )}
            </div>
        );
    }
}

function createCustomField<TProps, TValue>(
    targetProperty: string,
    key: string,
    component: React.ComponentType<TProps>,
    createProps: (onChange: (value: TValue) => void) => TProps
): IPropertyPaneField<ICustomFieldProperties> {
    return {
        type: PropertyPaneFieldType.Custom,
        targetProperty,
        properties: {
            key,
            onRender: (element, _context, changeCallback) => {
                if (!changeCallback) {
                    throw new Error(`Property pane change callback is missing for ${targetProperty}.`);
                }
                let root = propertyPaneRoots.get(element);
                if (!root) {
                    root = createRoot(element);
                    propertyPaneRoots.set(element, root);
                }
                const props = createProps(value => changeCallback(targetProperty, value));
                root.render(React.createElement(component, props));
            },
            onDispose: (element: HTMLElement) => {
                const root = propertyPaneRoots.get(element);
                if (root) {
                    root.unmount();
                    propertyPaneRoots.delete(element);
                }
            }
        }
    };
}

export function PropertyPaneCodeEditor(
    targetProperty: string,
    props: Omit<ICodeEditorProps, 'onSave'>
): IPropertyPaneField<ICustomFieldProperties> {
    return createCustomField<ICodeEditorProps, string>(
        targetProperty,
        'codeEditorFieldId',
        CodeEditor,
        onSave => ({ ...props, onSave })
    );
}

export function PropertyPaneAudiencePicker(
    targetProperty: string,
    context: WebPartContext,
    selected: IAudience[]
): IPropertyPaneField<ICustomFieldProperties> {
    return createCustomField<IPeoplePickerProps, IAudience[]>(
        targetProperty,
        'audienceTargeting',
        AudiencePicker,
        onChange => ({ context, label: 'Target Audience', selected, onChange })
    );
}

export function PropertyPaneNumberField(
    targetProperty: string,
    label: string,
    description: string,
    value: number
): IPropertyPaneField<ICustomFieldProperties> {
    return createCustomField<INumberFieldProps, number>(
        targetProperty,
        targetProperty,
        NumberField,
        onChange => ({ label, description, value, onChange })
    );
}
