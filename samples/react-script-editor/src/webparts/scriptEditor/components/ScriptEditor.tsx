import * as React from 'react';
import { IScriptEditorProps } from './IScriptEditorProps';
import styles from './ScriptEditor.module.scss';

export default class ScriptEditor extends React.Component<IScriptEditorProps, any> {
    constructor(props: IScriptEditorProps, state: any) {
        super(props);

        this._showDialog = this._showDialog.bind(this);
        this.state = state || {};
    }

    public componentDidMount(): void {
        this.setState({ script: this.props.script, loaded: this.props.script });
    }

    private _showDialog(): void {
        this.props.propPaneHandle.open();
    }

    public render(): React.ReactElement<IScriptEditorProps> {
        const viewMode = <span dangerouslySetInnerHTML={{ __html: this.state.script }} />;
        return (
            <div className='ms-Fabric'>
                <div className={styles.scriptEditorPlaceholder}>
                    <span className={styles.scriptEditorPlaceholderIcon} aria-hidden='true'>JS</span>
                    <div className={styles.scriptEditorPlaceholderContent}>
                        <h2 className={styles.scriptEditorPlaceholderTitle}>{this.props.title}</h2>
                        <p className={styles.scriptEditorPlaceholderDescription}>Please configure the web part</p>
                        <button className={styles.scriptEditorPlaceholderButton} type='button' onClick={this._showDialog}>
                            Edit markup
                        </button>
                    </div>
                </div>
                {viewMode}
            </div>);
    }
}
