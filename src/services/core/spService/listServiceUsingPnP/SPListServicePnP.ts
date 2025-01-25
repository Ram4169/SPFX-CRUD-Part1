import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from '@pnp/sp';
import { ISPListService } from './ISPListServicePnP';
import { getSP } from '../../../../configurations/pnpJSConfig/pnpJSConfig';

export class SPListService implements ISPListService {
  private _sp: SPFI;
  constructor(context: WebPartContext) {
    this._sp = getSP(context);
  }

  public async getListItems(listTitle: string): Promise<any> {
    const response = await this._sp.web.lists.getByTitle(listTitle).items();
    return response;
  }
}
