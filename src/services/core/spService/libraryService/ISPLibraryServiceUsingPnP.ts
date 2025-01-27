export interface ISPLibraryService {
  /**
   * Get SharePoint Library Files/Folder
   * @param libraryName SharePoint Library Display Name
   */
  getRootFolders(libraryName: string): Promise<any[]>;

  /**
   * This method returns all the files and folders from the document library recursively
   * @param libraryName
   */
  getAllItems(libraryName: string): Promise<any[]>;

  getFoldersUsingRelativePath(relativePath: string): Promise<any[]>;

  getFolderByServerRelativePathFromDifferentWeb(
    tenantUrl: string,
    relativePath: string
  ): Promise<any[]>;

  getFilesUsingRelativePath(relativePath: string): Promise<any[]>;

  getFilesByServerRelativePathFromDifferentWeb(
    tenantUrl: string,
    relativePath: string
  ): Promise<any[]>;
}
