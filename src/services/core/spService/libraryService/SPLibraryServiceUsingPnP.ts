import { SPFI } from '@pnp/sp';
import { ISPLibraryService } from './ISPLibraryServiceUsingPnP';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from '../../../../configurations/pnpJSConfig/pnpJSConfig';

export class SPLibraryService implements ISPLibraryService {
  private _sp: SPFI;
  constructor(context: WebPartContext) {
    this._sp = getSP(context);
  }

  public async getRootFolders(libraryName: string): Promise<any[]> {
    // gets list's folders
    const listFolders = await this._sp.web.lists
      .getByTitle(libraryName)
      .rootFolder.folders.filter('ListItemAllFields/Id ne null')
      .expand('ListItemAllFields')();
    return listFolders;
  }

  public async getAllItems(libraryName: string): Promise<any[]> {
    const _documentItems = await this._sp.web.lists
      .getByTitle(libraryName)
      .items.select('*')
      .expand('File/Length')();
    return _documentItems;
  }

  public async getFolderByServerRelativePath(
    relativePath: string
  ): Promise<any[]> {
    // folder is an IFolder and supports all the folder operations
    let folders = this._sp.web
      .getFolderByServerRelativePath(relativePath)
      .folders.filter('ListItemAllFields/Id ne null')
      .expand('Files/ListItemAllFields')
      .select('*')();

    return folders;
  }

  public async getFilesByServerRelativePath(
    relativePath: string
  ): Promise<any[]> {
    // folder is an IFolder and supports all the folder operations
    let files = this._sp.web
      .getFolderByServerRelativePath(relativePath)
      .files.filter('ListItemAllFields/Id ne null')
      .expand('ListItemAllFields')
      .select('*')();

    return files;
  }
}
