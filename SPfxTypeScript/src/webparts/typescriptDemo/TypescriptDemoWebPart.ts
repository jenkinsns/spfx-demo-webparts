import { employee } from './../Modules/FirstModule';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './TypescriptDemoWebPart.module.scss';
import * as strings from 'TypescriptDemoWebPartStrings';


export interface ITypescriptDemoWebPartProps {
  description: string;
}

export default class TypescriptDemoWebPart extends BaseClientSideWebPart <ITypescriptDemoWebPartProps> {

  public render(): void {
    let obj = new employee("Jenkins NS",120);

    let welcomemsg:string = obj.printfromfunction("Oliver", 350);

    this.domElement.innerHTML = `
      <div class="${ styles.typescriptDemo }">
    <div class="${ styles.container }">
      <div class="${ styles.row }">
        <div class="${ styles.column }">
          <span class="${ styles.title }">${welcomemsg} !</span>

          </a>
          </div>
          </div>
          </div>
          </div>`;
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
