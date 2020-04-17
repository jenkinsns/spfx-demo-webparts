import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './GetSPlistitemsnojsWebPart.module.scss';
import * as strings from 'GetSPlistitemsnojsWebPartStrings';

import {  SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions} from '@microsoft/sp-http';
import {  Environment,  EnvironmentType } from '@microsoft/sp-core-library';

import MockHttpClient from './MockHttpClient';

export interface IGetSPlistitemsnojsWebPartProps {
  description: string;
}

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  ContactNumber: string;
  CompanyName: string;
  Country: string;
}

export default class GetSPlistitemsnojsWebPart extends BaseClientSideWebPart <IGetSPlistitemsnojsWebPartProps> {

  private _getMockListData(): Promise<ISPLists> {
    return MockHttpClient.get(this.context.pageContext.web.absoluteUrl).then(() => {
        const listData: ISPLists = {
            value:
            [
               { Title: 'Mock Contact Person', ContactNumber: '9840462655', CompanyName: 'Jenkins',Country: 'India'},
               { Title: 'Mock Contact Person', ContactNumber: '9840462655', CompanyName: 'Jenkins',Country: 'India'},
               { Title: 'Mock Contact Person', ContactNumber: '9840462655', CompanyName: 'Jenkins',Country: 'India'},
            ]
            };

        return listData;
    }) as Promise<ISPLists>;
}

  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('Contactlist')/Items",SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
        return response.json();
        });
    }

  private _renderListAsync(): void {

    if (Environment.type === EnvironmentType.Local) {
      this._getMockListData().then((response) => {
        this._renderList(response.value);
      }); }
      else {
      this._getListData()
        .then((response) => {
          this._renderList(response.value);
        });
    }

   }

  private _renderList(items: ISPList[]): void {
      let html: string = `<table cellpadding="5" class="${styles.rtable}">
        <tr class="${styles.rrow}">
            <th clsss="${styles.rheader}">Contact Person</th>
            <th clsss="${styles.rheader}">Contact Number</th>
            <th clsss="${styles.rheader}">Company Name</th>
            <th clsss="${styles.rheader}">Country</th>
        </tr>
        `;
      items.forEach((item: ISPList) => {
        html += `
        <tr>
            <td clsss="${styles.rcell}">${item.Title}</td>
            <td clsss="${styles.rcell}">${item.ContactNumber}</td>
            <td clsss="${styles.rcell}">${item.CompanyName}</td>
            <td clsss="${styles.rcell}">${item.Country}</td>
        </tr>
        `;
      });
      html += '</table>';

      const listContainer: Element = this.domElement.querySelector('#spListItems');
      listContainer.innerHTML = html;
    }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.getSPlistitemsnojs }">
          <div id="spListItems" class="${ styles.container }"/>
      </div>`;
      this._renderListAsync();
  }

  protected get dataVersion(): Version {
  return Version.parse('1.0');
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
              })
            ]
          }
        ]
      }
    ]
  };
}
}
