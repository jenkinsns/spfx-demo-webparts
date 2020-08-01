import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './DyanmicdataloadWebPart.module.scss';
import * as strings from 'DyanmicdataloadWebPartStrings';
import {  SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions} from '@microsoft/sp-http';

export interface IDyanmicdataloadWebPartProps {
  description: string;
  checkboxProperty1:boolean;
  checkboxProperty2:boolean;
  dropdown1:string;
}

export interface ISPItems{
  value: ISPItem[];
}
export interface ISPItem{
  Title: string;
  Id:string;
}

export interface spListItems{
  value: spListItem[];
}
export interface spListItem{
  Title: string;
  id: string;
  Created: string;
  Author: {
    Title: string;
  };
}


export default class DyanmicdataloadWebPart extends BaseClientSideWebPart <IDyanmicdataloadWebPartProps> {

  private listName: string = "";
  private checkboxProperty1: string = "Created";
  private checkboxProperty2: string = "Author";
  private _options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();

  public render(): void {
    this.domElement.innerHTML = `<div class="${styles.dyanmicdataload}">
    <div class="${styles.Table}">
      <div class="${styles.Heading}">
        <div class="${styles.Cell}">Title</div>
      </div>
    </div>
  </div>`;
  this.listName = this.properties.description;
  this.LoadData();
  }

  private LoadData(): void{
    let url: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('"+this.listName+"')/items?$select=Title";
    // If Created Time check box option is selected
    if(this.properties.checkboxProperty1){
      url += ",Created";
      // Column header for Created Time field
       this.domElement.querySelector("."+styles.Heading).innerHTML +=`<div class="${styles.Cell}">Created</div>`;
    }
    // If Author check box option is selected
    if(this.properties.checkboxProperty2){
      url += ",Author/Title&$expand=Author";
      // Column header for Author field
      this.domElement.querySelector("."+styles.Heading).innerHTML +=`<div class="${styles.Cell}">Author</div>`;
    }
    this.GetListData(url).then((response)=>{
      // Render the data in the web part
      this.RenderListData(response.value);
    });
  }

  private GetListData(url: string): Promise<spListItems>{
    // Retrieves data from SP list
    return this.context.spHttpClient.get(url,SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse)=>{
       return response.json();
    });
  }
  private RenderListData(listItems: spListItem[]): void{
    let itemsHtml: string = "";
    // Displays the values in table rows
    if(listItems)
    {
    listItems.forEach((listItem: spListItem)=>{
      let itemTimeStr: string = listItem.Created;
      let itemTime: Date = new Date(itemTimeStr);
      itemsHtml += `<div class="${styles.Row}">`;
      itemsHtml += `<div class="${styles.Cell}"><p>${listItem.Title}</p></div>`;
      if(this.properties.checkboxProperty1){
        itemsHtml += `<div class="${styles.Cell}"><p>${listItem.Created}</p></div>`;
      }
      if(this.properties.checkboxProperty2){
        itemsHtml += `<div class="${styles.Cell}"><p>${listItem.Author.Title}</p></div>`;
      }

      itemsHtml += `</div>`;
    });
  }
    this.domElement.querySelector("."+styles.Table).innerHTML +=itemsHtml;
  }



  private getDataforDropdown()
  {
    let url: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('Country')/items";
    this.GetListDataforDorpdown(url).then((response)=>{
      this.RenderListDataforDorpdown(response.value);
    });
  }

  private GetListDataforDorpdown(url: string): Promise<ISPItems>{
    // Retrieves data from SP list
    return this.context.spHttpClient.get(url,SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse)=>{
       return response.json();
    });
  }

  private RenderListDataforDorpdown(listItems: ISPItem[]){
    let options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
    if(listItems)
    {
      listItems.forEach((listItem: ISPItem)=>{
      let _title: string = listItem.Title;
      let _id: string = listItem.Id;
      options.push({ key: _id, text: _title });
    });
    }
  this._options = options;
  }



  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    this.getDataforDropdown();
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
                label: "List Name",
                placeholder:"Enter your list name"
              })
            ]
          },
          {
            groupName: "OptionalFields",
            groupFields: [
              PropertyPaneCheckbox('checkboxProperty1',{
                checked:false,
                disabled:false,
                text: this.checkboxProperty1
              }),
              PropertyPaneCheckbox('checkboxProperty2',{
                checked:false,
                disabled:false,
                text: this.checkboxProperty2
              }),
              PropertyPaneDropdown('dropdown1',{
                label:"Dynamic DropDown",
                options:this._options
              })
            ]
          }
        ]
      }
    ]
  };
}
protected get dataVersion(): Version {
  return Version.parse('1.0');
  }

}
