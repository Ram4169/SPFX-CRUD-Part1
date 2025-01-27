import { spfi, SPFI, SPFx } from '@pnp/sp';
import { ISPLibraryService } from './ISPLibraryServiceUsingPnP';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from '../../../../configurations/pnpJSConfig/pnpJSConfig';
import { AssignFrom } from '@pnp/core';
import { Web } from '@pnp/sp/webs';

export class SPLibraryService implements ISPLibraryService {
  private _sp: SPFI;
  private _context: WebPartContext;
  constructor(context: WebPartContext) {
    this._sp = getSP(context);
    this._context = context;
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

  public async getFoldersUsingRelativePath(
    relativePath: string
  ): Promise<any[]> {
    // folder is an IFolder and supports all the folder operations
    let folders = this._sp.web
      .getFolderByServerRelativePath(relativePath)
      .folders.filter('ListItemAllFields/Id ne null')
      .expand('ListItemAllFields')
      .select('*')();

    return folders;
  }

  public async getFolderByServerRelativePathFromDifferentWeb(
    tenantUrl: string,
    relativePath: string
  ): Promise<any[]> {
    let relativePathSplit = relativePath.split('/');
    let siteUrl = tenantUrl + relativePathSplit[1] + '/' + relativePathSplit[2];
    // folder is an IFolder and supports all the folder operations
    const spWebB = spfi(siteUrl).using(SPFx(this._context));

    let folders = spWebB.web
      .getFolderByServerRelativePath(relativePath)
      .folders.filter('ListItemAllFields/Id ne null')
      .expand('ListItemAllFields')
      .select('*')();

    return folders;
  }

  public async getFilesUsingRelativePath(relativePath: string): Promise<any[]> {
    // folder is an IFolder and supports all the folder operations
    let files = this._sp.web
      .getFolderByServerRelativePath(relativePath)
      .files.filter('ListItemAllFields/Id ne null')
      .expand('ListItemAllFields')
      .select('*')();

    return files;
  }

  public async getFilesByServerRelativePathFromDifferentWeb(
    tenantUrl: string,
    relativePath: string
  ): Promise<any[]> {
    let relativePathSplit = relativePath.split('/');
    let siteUrl = tenantUrl + relativePathSplit[1] + '/' + relativePathSplit[2];
    // folder is an IFolder and supports all the folder operations
    //const spWebB = spfi(siteUrl).using(SPFx(this._context));
    //const spWebB = spfi(siteUrl).using(AssignFrom(this._sp.web));
    const spWebB = Web(siteUrl).using(AssignFrom(this._sp.web));

    let files = spWebB
      .getFolderByServerRelativePath(relativePath)
      .files.filter('ListItemAllFields/Id ne null')
      .expand('ListItemAllFields')
      .select('*')();

    return files;
  }
}
