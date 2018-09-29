import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { MSGraphClient } from '@microsoft/sp-http';
import {
    autobind,
    ColorPicker,
    PrimaryButton,
    Button,
    DialogFooter,
    DialogContent,
    CommandButton,
    Label,
    TextField
  } from 'office-ui-fabric-react';
import { WebPartContext } from '@microsoft/sp-webpart-base';

interface IAddFolderDialogContentProps{
    close: () => void;
    createFolder: (name: string, parent?:string) => void;
    parentFolder?:string;
}
interface IAddFolderDialogContentState{
    folderName:string;
}

class AddFolderDialogContent extends React.Component<IAddFolderDialogContentProps, IAddFolderDialogContentState>{
    constructor(props){
        super(props);
        this.state = {
            folderName: ""
        };
    }

    public render(): JSX.Element{
        return (
            <div>
                <DialogContent title="Add Folder" showCloseButton={true} onDismiss={ this.props.close }>
                <TextField label={"Folder Name:"}
                  required={true}
                  value={this.state.folderName}
                  onChanged={this._onNameChange}
                  onGetErrorMessage={this._getNameErrorMessage}
                />
                    <DialogFooter>
                        <CommandButton className="btn btn-primary"  text='Cancel' title='Cancel' onClick={this.props.close}  />
                        <PrimaryButton text='Create Folder' title='Create Folder' onClick={ () =>{ this.props.createFolder(this.state.folderName, this.props.parentFolder);} }></PrimaryButton>
                    </DialogFooter>
                </DialogContent>
            </div>

        );
    }

    @autobind
    private _onNameChange(newValue: string): void {
      this.setState({
        folderName: newValue,
      });
    }

    private _getNameErrorMessage(value: string): string {
        return (value !== null || value.length > 0)
          ? ''
          : `${"Value cannot be null or empty."}`;
      }

}

export default class AddFolderDialog extends BaseDialog {
    private _context:WebPartContext;
    private _parent:string;
    constructor(context:WebPartContext,parentFolderId?:string){
        super();
        this._context = context;
        this._parent = parentFolderId;
        this._createFolder = this._createFolder.bind(this);
    }
    protected render(): void {
        ReactDOM.render(<AddFolderDialogContent createFolder={this._createFolder} close={this.close} parentFolder={this._parent} />, this.domElement);
    }

    private _createFolder(folderName:string,parentFolderId?:string): void{
          this._context.msGraphClientFactory.getClient().then((graphClient:MSGraphClient):void =>{
            let query:string;
            if(parentFolderId){
              query = `/me/drive/items/${parentFolderId}/children`;
            }else{
              query = "/me/drive/root/children";
            }
            graphClient.api(query).version("v1.0").post({"name":folderName,"folder": { },"@microsoft.graph.conflictBehavior": "rename"},(error, result)=>{
              if(error){
                console.error(error);
              }
              this.close();
            });
          });
    
      }

}