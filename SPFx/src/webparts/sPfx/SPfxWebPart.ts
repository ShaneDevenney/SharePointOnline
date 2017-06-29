import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  IWebPartContext
} from '@microsoft/sp-webpart-base';

import * as strings from 'sPfxStrings';
import SPfx from './components/SPfx';
import { ISPfxProps } from './components/ISPfxProps';
import { ISPfxWebPartProps } from './ISPfxWebPartProps';
import MockHttpClient from './MockHttpClient';
import styles from './components/SPfx.module.scss';

import {
    SPHttpClient,
    SPHttpClientResponse
} from '@microsoft/sp-http';

import {
    Environment,
    EnvironmentType
} from '@microsoft/sp-core-library';

export interface ISPLists {
    value: ISPList[];
}

export interface ISPList {
    Title: string;
    Id: string;
}

export default class SPfxWebPart extends BaseClientSideWebPart<ISPfxWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISPfxProps > = React.createElement(
      SPfx,
      {
          description: this.properties.description,
          webname: this.context.pageContext.web.title
      }
    );

    ReactDom.render(element, this.domElement);
    this._renderListAsync();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private _getMockListData(): Promise<ISPLists> {
      return MockHttpClient.get()
          .then((data: ISPList[]) => {
              var listData: ISPLists = { value: data };
              return listData;
          }) as Promise<ISPLists>;
  }

  private _getListData(): Promise<ISPLists> {
      return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=Hidden eq false`, SPHttpClient.configurations.v1)
          .then((response: SPHttpClientResponse) => {
              return response.json();
          });
  }

  private _renderListAsync(): void {
      // Local environment
      if (Environment.type === EnvironmentType.Local) {
          this._getMockListData().then((response) => {
              this._renderList(response.value);
          });
      }
      else if (Environment.type == EnvironmentType.SharePoint ||
          Environment.type == EnvironmentType.ClassicSharePoint) {
          this._getListData()
              .then((response) => {
                  this._renderList(response.value);
              });
      }
  }

  private _renderList(items: ISPList[]): void {
      let html: string = '';
      items.forEach((item: ISPList) => {
          html += `
        <ul class="${styles.list}">
            <li class="${styles.listItem}">
                <span class="ms-font-l">${item.Title}</span>
            </li>
        </ul>`;
      });

      const listContainer: Element = this.domElement.querySelector('#spListContainer');
      listContainer.innerHTML = html;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                  }),
                PropertyPaneTextField('test', {
                    label: 'Multi-line Text Field',
                    multiline: true
                }),
                PropertyPaneCheckbox('test1', {
                    text: 'Checkbox'
                }),
                PropertyPaneDropdown('test2', {
                    label: 'Dropdown',
                    options: [
                        { key: '1', text: 'One' },
                        { key: '2', text: 'Two' },
                        { key: '3', text: 'Three' },
                        { key: '4', text: 'Four' }
                    ]
                }),
                PropertyPaneToggle('test3', {
                    label: 'Toggle',
                    onText: 'On',
                    offText: 'Off'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
