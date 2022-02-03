import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import {
    ColorPicker,
    PrimaryButton,
    Button,
    DialogFooter,
    DialogContent,
    IColor
} from 'office-ui-fabric-react';

interface IShowItemDetailsContentProps {
    message: string;
    htmlElements: any;
    close: () => void;
    submit: (color: string) => void;
    defaultColor?: string;
}

class ShowItemDetailsContent extends React.Component<IShowItemDetailsContentProps, {}> {

    public render(): JSX.Element {
        var iHtml = this.props.htmlElements[3].textContent;

       function createMarkup() 
       {
        return {__html: iHtml};
       }
       
       return <DialogContent
            title=''
            subText=''
            onDismiss={this.props.close}
            showCloseButton={true}

        >
            <div>
                <h2>{this.props.htmlElements[0].textContent}</h2>
                <p><b>Saksnummer:</b> <span>{this.props.htmlElements[4].textContent}</span></p>
                <p><b>Start:</b> <span>{this.props.htmlElements[1].textContent}</span></p>
                <p><b>Description:</b></p>
                <div dangerouslySetInnerHTML={createMarkup()} />
            </div>
            <DialogFooter>
                <Button text='Cancel' title='Cancel' onClick={this.props.close} />
                <PrimaryButton text='OK' title='OK' onClick={() => { this.props.submit(""); }} />
            </DialogFooter>
        </DialogContent>;
    }
}

export default class ShowItemDetails extends BaseDialog {
    public message: string;
    public htmlElements: any;
    //public colorCode: string;

    public render(): void {
        ReactDOM.render(<ShowItemDetailsContent
            close={this.close}
            message={this.message}
            submit={this._submit}
            htmlElements={this.htmlElements}
        />, this.domElement);
    }

    public getConfig(): IDialogConfiguration {
        return {
            isBlocking: false
        };
    }

    protected onAfterClose(): void {
        super.onAfterClose();

        // Clean up the element for the next dialog
        ReactDOM.unmountComponentAtNode(this.domElement);
    }

    private _submit = (color: string) => {
        //this.colorCode = color;
        this.close();
    }
}