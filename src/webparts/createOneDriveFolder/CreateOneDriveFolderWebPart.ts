import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'CreateOneDriveFolderWebPartStrings';
import CreateOneDriveFolder from './components/CreateOneDriveFolder';
import { ICreateOneDriveFolderProps } from './components/ICreateOneDriveFolderProps';
import { MSGraphClient } from '@microsoft/sp-http';
import { IFolder } from '../../common/IFolder';

export interface ICreateOneDriveFolderWebPartProps {
  description: string;
}

export default class CreateOneDriveFolderWebPart extends BaseClientSideWebPart<ICreateOneDriveFolderWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ICreateOneDriveFolderProps > = React.createElement(
      CreateOneDriveFolder,
      {
        description: this.properties.description,
        context:this.context,
        loadOneDriveItems: this._loadItemsFromOneDrive.bind(this),
        createFolder: this._createFolder.bind(this)
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private _createFolder(folderName:string,parentFolderId?:string): Promise<IFolder>{
    return new Promise<IFolder>((resolve:(folder:IFolder)=> void,reject:(error:any) => void) =>{
      this.context.msGraphClientFactory.getClient().then((graphClient:MSGraphClient):void =>{
        let query:string;
        if(parentFolderId){
          query = `/me/drive/items/${parentFolderId}/children`;
        }else{
          query = "/me/drive/root/children";
        }
        graphClient.api(query).version("v1.0").post({"name":folderName,"folder": { },"@microsoft.graph.conflictBehavior": "rename"},(error, result)=>{
          if(error){
            console.error(error);
            reject(error);
          }
          resolve({Id:result.id,Name:result.name});
        });
      });

    });
  }

  private _loadItemsFromOneDrive(id?:string): Promise<any>{
    return new Promise<any>((resolve: (options: any) => void, reject: (error: any) => void) => {
      this.context.msGraphClientFactory.getClient().then((graphClient: MSGraphClient):void =>{
        let query:string;
        if(id){
          query = `me/drive/items/${id}/children`;
        }else{
          query = "me/drive/root/children";
        }
        graphClient.api(query).version("v1.0")/*.select("displayName,mail,id")*/.get((error, result) => {
          if (error) {
            console.error(error);
            reject(error);
          }
          // Prepare the output array
          var items: Array<any> = new Array<any>();
          // Map the JSON response to the output array
          result.value.map((item: any) => {
            if(item.folder){
               items.push({
                 icon: "",
              Name: item.name,
              Id: item.id,
              folder: item.folder
            });
            }
           
  
          });
          resolve(items);
        });

      });
      
    });
  }
}
