import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { Web, IWeb } from "@pnp/sp/webs";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IItemAddResult } from "@pnp/sp/items";
import './notify.css';
import {
    PrimaryButton,
    Button,
    DialogFooter,
    DialogContent,
} from '@fluentui/react';
import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';
import ClassicEditor from 'ckeditor5-classic';

ClassicEditor.defaultConfig = {    
    toolbar: {    
      items: [    
        'heading',    
        '|',    
        'bold',    
        'italic',    
        'fontSize',    
        'fontFamily',    
        'fontColor',    
        'fontBackgroundColor',    
       // 'link',    
        'bulletedList',    
        'numberedList',    
       // 'imageUpload',    
      //  'insertTable',    
        'blockQuote',    
        'undo',    
        'redo'    
      ]    
    },    
    image: {    
      toolbar: [    
        'imageStyle:full',    
        'imageStyle:side',    
        '|',    
        'imageTextAlternative'    
      ]    
    },    
    fontFamily: {    
      options: [    
        'Arial',    
        'Helvetica, sans-serif',    
        'Courier New, Courier, monospace',    
        'Georgia, serif',    
        'Lucida Sans Unicode, Lucida Grande, sans-serif',    
        'Tahoma, Geneva, sans-serif',    
        'Times New Roman, Times, serif',    
        'Trebuchet MS, Helvetica, sans-serif',    
        'Verdana, Geneva, sans-serif'    
      ]    
    },    
    language: 'en'    
  };

interface IAddItemDetailsContentProps {
    close: () => void;
    ctx: ApplicationCustomizerContext;
    submit: (ctx: ApplicationCustomizerContext, e) => void;
}

class AddItemDetailsContent extends React.Component<IAddItemDetailsContentProps, {}> {

    public render(): JSX.Element {
        return <DialogContent
            title='Add new notification'
            subText=''
            
            onDismiss={this.props.close}
            showCloseButton={true}
            
           
        >
            <div id="divContainer">
                <table style={{width:'100%'}}>
                    <tr>
                        <td className='tagLine'>Title</td>
                        <td>
                            <input type='text' className='TypesInput formControl' id='txtTitle' style={{width:'100%'}}></input>
                        </td>
                    </tr>
                    <tr>
                        <td className='tagLine'>Description</td>
                        <td>
                            <textarea id='txtBody'></textarea>
                        </td>
                    </tr>
                    <tr>
                        <td className='tagLine'>Start Date</td>
                        <td>
                            <input className='TypesInput' type='datetime-local' id='txtStart'></input>
                        </td>
                    </tr>
                    <tr>
                        <td className='tagLine'>End date</td>
                        <td>
                            <input className='TypesInput' type='datetime-local' id='txtEnd'></input>
                        </td>
                    </tr>
                    {/* <tr>
                        <td className='tagLine'>Sales No</td>
                        <td>
                            <input className='TypesInput' type='text' id='txtSale'></input>
                        </td>
                    </tr> */}
                </table>
            </div>
            <DialogFooter>
                <Button text='Cancel' title='Cancel' onClick={this.props.close} />
                <PrimaryButton text='OK' title='OK' onClick={(e) => { this.props.submit(this.props.ctx, e) }} />
            </DialogFooter>
        </DialogContent>;
    }
}

export default class AddItemDetails extends BaseDialog {
    public editorBody;
    public ctx: ApplicationCustomizerContext;
    public render(): void {
        ReactDOM.render(<AddItemDetailsContent
            close={this.close}
            submit={this._submit}
            ctx={this.ctx}
        />, this.domElement);

        this._initializeCKeditor(this.domElement);
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
        window.location.href = window.location.href;
    }


    private _submit = (ctx: ApplicationCustomizerContext, e) => {
        //this.colorCode = color;
        var adminUrl = ctx.pageContext.web.absoluteUrl.replace(ctx.pageContext.web.serverRelativeUrl, "");
        adminUrl = adminUrl + "/sites/valoadmin";

        // Provision SP list if it's not available already. Will be executed while adding the extension to menu bar.
        // User needs to have edit rights in the site for list provision.
        var adminWeb = Web(adminUrl);
        adminWeb.lists.getByTitle("NotificationConfig").
            items.select("Title").top(1).
            orderBy("Modified", true).
            get().
            then((items: any) => {
                if (items.length > 0) {
                    var siteUrl = items[0].Title;
                    var siteWeb = Web(siteUrl);

                    // add an item to the list
                    const lastItem = siteWeb.lists.getByTitle("NotificationList").items.add({
                        Title: (document.getElementById("txtTitle") as HTMLInputElement).value,
                        Description: this.editorBody.getData(), //(document.getElementById("txtDesc") as HTMLInputElement).value,
                        StartDate: (document.getElementById("txtStart") as HTMLInputElement).value,
                        EndDate: (document.getElementById("txtEnd") as HTMLInputElement).value
                      //  SalesNo: (document.getElementById("txtSale") as HTMLInputElement).value
                    });

                    lastItem.then(i => {
                        alert("Item saved successfully !");
                        window.location.href = window.location.href;
                    });
                }
            });

        // this.close();
    }

      /* Load CKeditor RTE*/    
      public _initializeCKeditor(htmlEl: HTMLElement): void {    
        try {    
        /*Replace textarea with classic editor*/    
        ClassicEditor    
            .create(htmlEl.querySelector("#txtBody"), {    
            }).then(editor => {    
                this.editorBody = editor;
                console.log("CKEditor5 initiated");  
            }).catch(error => {    
            console.log("Error in Classic Editor Create " + error);    
            });    
        } catch (error) {    
        console.log("Error in  InitializeCKeditor " + error);    
        }    
    }
}