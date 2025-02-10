import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { ISPListServiceSPHTTP } from './ISPListServiceSPHTTP';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export class SPListServiceSPHTTP implements ISPListServiceSPHTTP {
  private _context: WebPartContext;
  constructor(context: WebPartContext) {
    this._context = context;
  }

  public getListItems = (listTitle: string): Promise<any> => {
    return new Promise((resolve, reject) => {
      this._context.spHttpClient
        .get(
          this._context.pageContext.web.absoluteUrl +
            `/_api/web/lists/getbytitle('${listTitle}')/items`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: 'application/json;odata=nometadata',
              'odata-version': '',
            },
          }
        )
        .then((response: SPHttpClientResponse) => {
          if (response.ok) {
            response.json().then((responseJSON) => {
              resolve(responseJSON.value);
            });
          } else {
            response.json().then((responseJSON) => {
              console.log(responseJSON);
              alert(
                `Something went wrong! Check the error in the browser console.`
              );
              resolve([]);
            });
          }
        })
        .catch((error) => {
          console.log(error);
          resolve([]);
        });
    });

    //return result;
  };
}
