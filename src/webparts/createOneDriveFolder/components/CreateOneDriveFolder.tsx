import * as React from 'react';
import styles from './CreateOneDriveFolder.module.scss';
import { ICreateOneDriveFolderProps } from './ICreateOneDriveFolderProps';
import {MarqueeSelection, DetailsList, Selection, Image, ImageFit, Link, CheckboxVisibility, IColumn, CommandBar, ICommandBar, ICommandBarProps, IContextualMenuItem} from 'office-ui-fabric-react';
import { escape } from '@microsoft/sp-lodash-subset';
import { ICreateOneDriveFolderState } from './ICreateOneDriveFolderState';
import { IFolder } from '../../../common/IFolder';
import AddFolderDialog from './AddFolderDialog';

export default class CreateOneDriveFolder extends React.Component<ICreateOneDriveFolderProps, ICreateOneDriveFolderState> {
  private _selection: Selection;
  private _columns: IColumn[];
  public constructor(props:ICreateOneDriveFolderProps){
    super(props);

    this._columns = [
      {
        key: 'FileIcon',
        name: '',
        headerClassName: 'DetailsListExample-header--FileIcon',
        className: 'DetailsListExample-cell--FileIcon',
        iconClassName: 'DetailsListExample-Header-FileTypeIcon',
        ariaLabel: 'Column operations for File type',
        iconName: 'Page',
        isIconOnly: true,
        fieldName: 'FileIcon',
        minWidth: 20,
        maxWidth: 20,
       // onColumnClick: this._onColumnClick,
       /* onRender: (item: any) => {
          return <img src={item.iconName} className={'DetailsListExample-documentIconImage'} />;
        }*/
      },
      {
        key: 'Name',
        name: 'Name',
        fieldName: 'Name',
        minWidth: 210,
        maxWidth: 350,
        isRowHeader: true,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
       // onColumnClick: this._onColumnClick,
        data: 'string',
        isPadded: true
      }
    ];
    this.state = {
      folders: [],
      history:[]
    };
    this._selection = new Selection;


    this._renderItemColumn = this._renderItemColumn.bind(this);
    this._folderClicked = this._folderClicked.bind(this);
    this._createFolder = this._createFolder.bind(this);
  }
  public componentDidMount(){
    this.props.loadOneDriveItems().then((items)=>{
      this.setState({folders:items});
    });
  }
  
  public render(): React.ReactElement<ICreateOneDriveFolderProps> {
    return (
      <div>
        <div className={styles.topBar}>
        <span className={styles.topBarText}>{this.props.description}</span>
          <CommandBar items={this._getCommandBarItems()} />
        </div>
        <div>
          <MarqueeSelection selection={this._selection}>
            <DetailsList columns={this._columns} items={this.state.folders} checkboxVisibility={CheckboxVisibility.hidden} onRenderItemColumn={this._renderItemColumn} />
          </MarqueeSelection>
        </div>
      </div>
    );
  }

  private _renderItemColumn(folder, index, column) {
    //here we can add column specific logic
    // - image control for the FileIcon column
    // - render link for the Name column
    let fieldContent = folder[column.fieldName];
    switch (column.key) {
      case 'FileIcon':
        return <Image src={"https://spoprod-a.akamaihd.net/files/odsp-next-prod_2018-09-07_20180919.004/odsp-media/images/itemtypes/20/folder.svg"} width={16} height={16} imageFit={ImageFit.center} />;
      case 'Name':
        return <Link data-selection-invoke={true} onClick={() =>{this._folderClicked(folder);}} >{folder[column.fieldName]}</Link>;
     // default:
       // return <Image src={"https://spoprod-a.akamaihd.net/files/odsp-next-prod_2018-09-07_20180919.004/odsp-media/images/itemtypes/20/folder.svg"} width={16} height={16} imageFit={ImageFit.center} />;
    }

  }

  private _folderClicked(folder){
    console.log(folder);
    this.props.loadOneDriveItems(folder["Id"]).then((folders:IFolder[])=>{
      let temp = this.state.history;
      temp.push(folder);
      this.setState({folders:folders, history:temp,currentFolder:folder});
    });
  }

  private _createFolder(){
    let folders: IFolder[] = this.state.folders;
    let dialog: AddFolderDialog ;

    if(this.state.currentFolder){
      dialog = new AddFolderDialog(this.props.context, this.state.currentFolder.Id);
      /*this.props.createFolder("Obaid",this.state.currentFolder.Id).then((folder:IFolder)=>{
       
        debugger;
        folders.push(folder);
          this.setState({folders:folders});
      });*/
    }else{
      dialog = new AddFolderDialog(this.props.context);
      /*this.props.createFolder("Obaid").then((folder:IFolder)=>{
        debugger;
        folders.push(folder);
          this.setState({folders:folders});
      });*/
    }
    dialog.show().then(() => {
    if(this.state.currentFolder){
      this._folderClicked(this.state.currentFolder);
    }else{
      this.props.loadOneDriveItems().then((items)=>{
        this.setState({folders:items});
      });
    }
      
    }).catch((error) => {
      console.log(error);
    });
  }

  private _getCommandBarItems(){

    const newIconButton: IContextualMenuItem = {
      key: 'newItem',
      name: 'New',
      cacheKey: 'myCacheKey', // changing this key will invalidate this items cache
      iconProps: {
        iconName: 'Add'
      },
      ariaLabel: 'New. Use left and right arrow keys to navigate',
      onClick: () =>{this._createFolder();}
    };
    
    return [
      
      newIconButton
    ];
  }
}
