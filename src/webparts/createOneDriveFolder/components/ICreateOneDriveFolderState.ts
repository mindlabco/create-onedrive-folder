import { IFolder } from "../../../common/IFolder";

export interface ICreateOneDriveFolderState {
    folders: IFolder[];
    history: IFolder[];
    currentFolder?: IFolder;
  //  selectionDetails: string;
  //isModalSelection: boolean;
}