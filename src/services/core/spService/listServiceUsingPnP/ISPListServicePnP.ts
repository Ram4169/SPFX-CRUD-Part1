export interface ISPListService {
  /**
   * Get SharePoint List Items
   * @param listTitle SharePoint List Display Name
   */
  getListItems(listTitle: string): Promise<any>;
}
