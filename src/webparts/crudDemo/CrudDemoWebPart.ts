import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'CrudDemoWebPartStrings';
import CrudDemo from './components/CrudDemo';
import { ICrudDemoProps } from './components/ICrudDemoProps';
import { getSP } from '../../configurations/pnpJSConfig/pnpJSConfig';
import { LibraryOption } from '../../common/LibraryOption';

export interface ICrudDemoWebPartProps {
  description: string;
  splibraryoption: string;
}

export default class CrudDemoWebPart extends BaseClientSideWebPart<ICrudDemoWebPartProps> {
  private _dropdownLibraryOptions: IPropertyPaneDropdownOption[] = [
    {
      key: LibraryOption['SP Service Using PnP'],
      text: LibraryOption[LibraryOption['SP Service Using PnP']],
    },
    {
      key: LibraryOption['SP Service Using SPHttp'],
      text: LibraryOption[LibraryOption['SP Service Using SPHttp']],
    },
    {
      key: LibraryOption['SP Service Using MSGraph Client'],
      text: LibraryOption[LibraryOption['SP Service Using MSGraph Client']],
    },
    {
      key: LibraryOption['SP Service Using AadHttp Client'],
      text: LibraryOption[LibraryOption['SP Service Using AadHttp Client']],
    },
  ];

  public render(): void {
    const element: React.ReactElement<ICrudDemoProps> = React.createElement(
      CrudDemo,
      {
        description: this.properties.description,
        context: this.context,
        selectedLibrary: Number(this.properties.splibraryoption),
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
                PropertyPaneDropdown('splibraryoption', {
                  label: 'Select a library',
                  options: this._dropdownLibraryOptions,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
