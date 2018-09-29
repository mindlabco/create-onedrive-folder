import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IFolder } from "../../../common/IFolder";

export interface ICreateOneDriveFolderProps {
  description: string;
  context: WebPartContext;
  loadOneDriveItems: (id?: string) => Promise<any[]>;
  createFolder: (folderName:string, id?: string) => Promise<IFolder>;
}
