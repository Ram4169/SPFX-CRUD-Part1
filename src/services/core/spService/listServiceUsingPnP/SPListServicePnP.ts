import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from '@pnp/sp';
import { ISPListService } from './ISPListServicePnP';
import { getSP } from '../../../../configurations/pnpJSConfig/pnpJSConfig';
import IEmployeeDetails from '../../../../models/IEmplyeeDetails';

export class SPListService implements ISPListService {
  private _sp: SPFI;
  constructor(context: WebPartContext) {
    this._sp = getSP(context);
  }

  public async getListItems(listTitle: string): Promise<IEmployeeDetails[]> {
    const response = await this._sp.web.lists.getByTitle(listTitle).items();
    return response;
  }

  public async getListItemById(
    listTitle: string,
    selectedColumn: string,
    expand: string,
    itemId: number
  ): Promise<IEmployeeDetails> {
    const response = await this._sp.web.lists
      .getByTitle(listTitle)
      .items.select(selectedColumn)
      .expand(expand)
      .getById(itemId)();
    return response;
  }

  public async createListItem(
    listTitle: string,
    data: IEmployeeDetails
  ): Promise<IEmployeeDetails> {
    const response = await this._sp.web.lists
      .getByTitle(listTitle)
      .items.add(data);
    return response;
  }

  public updateListItem(
    listTitle: string,
    id: number,
    data: IEmployeeDetails
  ): Promise<IEmployeeDetails> {
    const list = this._sp.web.lists.getByTitle(listTitle);
    return list.items.getById(id).update(data);
  }
}
