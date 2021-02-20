/* 
   https://bksdevsite.sharepoint.com/_layouts/15/workbench.aspx 
*/
//#region [imports]
  
  import * as React 
    from 'react';
  import * as ReactDom 
    from 'react-dom';
  import * as strings 
    from 'FaqSpFxClientSideWebPartWebPartStrings';
  import { Version } 
    from '@microsoft/sp-core-library';
  import { IPropertyPaneConfiguration, PropertyPaneTextField} 
    from '@microsoft/sp-property-pane';
  import {SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration} 
    from '@microsoft/sp-http';
  import { BaseClientSideWebPart } 
    from '@microsoft/sp-webpart-base';
  import {Environment,EnvironmentType} 
    from '@microsoft/sp-core-library';    
  import FaqSpFxClientSideWebPart 
    from './components/FaqSpFxClientSideWebPart';
  import {IFaqSpFxClientSideWebPartProps} 
    from './components/IFaqSpFxClientSideWebPartProps';
  import styles 
    from './components/FaqSpFxClientSideWebPart.module.scss';

 //#endregion

//#region [interfaces]
  
  export interface IFaqSpFxClientSideWebPartWebPartProps {
    title: string; 
    description: string;
    siteURL: string;
    list: string; 
  }

  export interface ISPLists {
    value: ISPList[];
  }
  
  export interface ISPList {
    Id: number;
    Title: string; 
    Answers: string;
    categor: string; 
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
      this._renderFAQItemsAsync();
    }
  
    private _renderFAQs(items: ISPList[]): void {
      const FAQContainer: HTMLElement = document.getElementById("FAQs"); 
      var htmlout : string = "";
      items.forEach((item: ISPList) => {
        htmlout += `<div class="${styles.row}">
                      <div class="${styles.question}">${item.Title}</div>
                      <div class="${styles.answer}">${item.Answers}</div>
                    </div>`;
      });
      FAQContainer.innerHTML = htmlout;
    }

  //#endregion

  //#region [AsyncCode]

    private _renderFAQItemsAsync(): void {
      if (Environment.type == EnvironmentType.SharePoint ||  
          Environment.type == EnvironmentType.ClassicSharePoint) {
          this._getFAQData()
            .then((response) => {
              this._renderFAQs(response.value);
            });
      } 
    }

  //#endregion

  //#region [QueryData]

    private _getFAQData(): Promise<ISPLists> {
      let restQuery = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Frequently Asked Questions')/items?
      &$select=Id,Title,Answers,Category`;
      console.log(restQuery);
      return this.context.spHttpClient.get(restQuery ,SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          return response.json();
        });   
    }

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