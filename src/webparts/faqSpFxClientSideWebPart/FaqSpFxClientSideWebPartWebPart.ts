/* 
   https://bksdevsite.sharepoint.com/_layouts/15/workbench.aspx 
*/
//#region [imports]
  
  import * as React from 'react';
  import * as ReactDom from 'react-dom';
  import { Version } from '@microsoft/sp-core-library';
  import { IPropertyPaneConfiguration, PropertyPaneTextField} from '@microsoft/sp-property-pane';
  import {SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration} from '@microsoft/sp-http';
  import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
  import * as strings from 'FaqSpFxClientSideWebPartWebPartStrings';
  import FaqSpFxClientSideWebPart from './components/FaqSpFxClientSideWebPart';
  import { IFaqSpFxClientSideWebPartProps } from './components/IFaqSpFxClientSideWebPartProps';

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

  //#region [DisplayCode]   
  
    public render(): void {
      const element: React.ReactElement<IFaqSpFxClientSideWebPartProps> = React.createElement(
        FaqSpFxClientSideWebPart,
        {
          title: this.properties.title, 
          description: this.properties.description
        }
      );
      ReactDom.render(element, this.domElement);
    }

  //#endregion

  //#region [AsyncCode]
  //#endregion

  //#region [QueryData]
  //#endregion

  //#region [GenericCode]

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
          },          
          {
            header: {
              description: strings.DataConnectionDescription
            },
            groups: [
              {
                groupName: strings.DataConnectionGroupName,
                groupFields: [
                  PropertyPaneTextField('Site URL', {
                    label: strings.SiteURLLabel
                  }),
                  PropertyPaneTextField('List Name', {
                    label: strings.ListLabel
                  })
                ]
              }
            ]
          }

        ]
      };
    }

  //#endregion

}