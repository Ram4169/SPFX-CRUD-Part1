import IEmployeeDetails from '../../../../models/IEmplyeeDetails';

export interface ISPListService {
  /**
   * Get list item by Id
   * @param listTitle List Title
   * @param selectedColumn Column names by comma seprated
   * @param expand lookup column names by comma separated
   * @param itemId list item id to be fetched
   */
  getListItemById(
    listTitle: string,
    selectedColumn: string,
    expand: string,
    itemId: number
  ): Promise<IEmployeeDetails>;

  /**
   * Get SharePoint List Items
   * @param listTitle SharePoint List Display Name
   */
  getListItems(listTitle: string): Promise<IEmployeeDetails[]>;

  /**
   * Creates list item
   * @param listTitle
   * @param data  payload
   * @returns Promise<any>
   */
  createListItem(
    listTitle: string,
    data: IEmployeeDetails
  ): Promise<IEmployeeDetails>;

  /**
   * Update list item
   * @param listTitle
   * @param id of the list item
   * @param data  payload
   * @returns Promise<any>
   */
  updateListItem(
    listTitle: string,
    id: number,
    data: IEmployeeDetails
  ): Promise<IEmployeeDetails>;
}
