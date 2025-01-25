import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'CrudDemoWebPartStrings';
import CrudDemo from './components/CrudDemo';
import { ICrudDemoProps } from './components/ICrudDemoProps';
import { getSP } from '../../configurations/pnpJSConfig/pnpJSConfig';

export interface ICrudDemoWebPartProps {
  description: string;
}

export default class CrudDemoWebPart extends BaseClientSideWebPart<ICrudDemoWebPartProps> {
  public render(): void {
    const element: React.ReactElement<ICrudDemoProps> = React.createElement(
      CrudDemo,
      {
        description: this.properties.description,
        context: this.context,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
    getSP(this.context);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
