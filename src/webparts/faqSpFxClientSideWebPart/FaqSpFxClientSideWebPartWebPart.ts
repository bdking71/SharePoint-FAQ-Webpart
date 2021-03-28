/* 
  /_layouts/15/workbench.aspx 

  ICON Reference: 
  https://thechriskent.com/2017/06/19/sharepoint-framework-app-icon/
  https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/basics/configure-web-part-icon

*/

//#region [imports]
  
  import * as React from 'react';
  import * as ReactDom from 'react-dom';
  import * as strings from 'FaqSpFxClientSideWebPartWebPartStrings';
  import { Version } from '@microsoft/sp-core-library';
  import { IPropertyPaneConfiguration, PropertyPaneTextField} from '@microsoft/sp-property-pane';
  import {SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration} from '@microsoft/sp-http';
  import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
  import {Environment,EnvironmentType} from '@microsoft/sp-core-library';    
  import FaqSpFxClientSideWebPart from './components/FaqSpFxClientSideWebPart';
  import {IFaqSpFxClientSideWebPartProps} from './components/IFaqSpFxClientSideWebPartProps';
  import styles from './components/FaqSpFxClientSideWebPart.module.scss';
  import { SPComponentLoader } from '@microsoft/sp-loader';
  import * as jQuery from 'jquery';
  import 'jqueryui'; 

  //#endregion

//#region [interfaces]
  
  export interface IFaqSpFxClientSideWebPartWebPartProps {
    title: string; 
    description: string;
    siteURL: string;
    list: string; 
  }

//#endregion  

export default class FaqSpFxClientSideWebPartWebPart extends BaseClientSideWebPart<IFaqSpFxClientSideWebPartWebPartProps> {
  
  public render(): void {
    const element: React.ReactElement<IFaqSpFxClientSideWebPartProps> = 
      React.createElement(FaqSpFxClientSideWebPart,
      {
        title: this.properties.title, 
        description: this.properties.description,
        siteURL: this.properties.siteURL,
        listName: this.properties.list,
        x: this
      }
    );
    ReactDom.render(element, this.domElement);
  }
  
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected getdataVersion(): Version {
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
                PropertyPaneTextField('title', {
                  label: strings.DescriptionFieldLabel
                }),
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