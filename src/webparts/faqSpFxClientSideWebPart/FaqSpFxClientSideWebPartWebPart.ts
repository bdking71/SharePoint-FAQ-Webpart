/* 
  https://bksdevsite.sharepoint.com/_layouts/15/workbench.aspx 

  npm install jquery@2
  npm install jqueryui

  npm install @types/jquery@2 --save-dev
  npm install @types/jqueryui --save-dev

*/

//#region [imports]
  
  import * as React from 'react';
  import * as ReactDom from 'react-dom';
  import * as strings from 'FaqSpFxClientSideWebPartWebPartStrings';
  import * as jQuery from 'jquery';
  import 'jqueryui';
  import { Version } from '@microsoft/sp-core-library';
  import { IPropertyPaneConfiguration, PropertyPaneTextField} from '@microsoft/sp-property-pane';
  import {SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration} from '@microsoft/sp-http';
  import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
  import {Environment,EnvironmentType} from '@microsoft/sp-core-library';    
  import FaqSpFxClientSideWebPart from './components/FaqSpFxClientSideWebPart';
  import {IFaqSpFxClientSideWebPartProps} from './components/IFaqSpFxClientSideWebPartProps';
  import styles from './components/FaqSpFxClientSideWebPart.module.scss';
  import { SPComponentLoader } from '@microsoft/sp-loader';

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
  }

//#endregion  

export default class FaqSpFxClientSideWebPartWebPart extends BaseClientSideWebPart<IFaqSpFxClientSideWebPartWebPartProps> {

  //#region [constructor]  

    public constructor() {
      super();
      SPComponentLoader.loadCss('https://code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css');
    }
  
    //#endregion  

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
      var htmlout : string = `<div class="accordion">`;
      items.forEach((item: ISPList) => {
        htmlout += `<h3>${item.Title}</h2>
                      <div>
                        <p>${item.Answers}</p>
                      </div>`;           
      });
      htmlout += `</div>`;
      FAQContainer.innerHTML = htmlout;
      
      const accordionOptions: JQueryUI.AccordionOptions = {
        animate: true,
        collapsible: false,
        icons: {
          header: 'ui-icon-circle-arrow-e',
          activeHeader: 'ui-icon-circle-arrow-s'
        }
      };

      jQuery('.accordion', this.domElement).accordion(accordionOptions);
    }

  //#endregion

  //#region [AsyncCode]

    private _renderFAQItemsAsync(): void {
      this._getFAQData()
        .then((response) => {
          this._renderFAQs(response.value);
        });
    }

  //#endregion

  //#region [QueryData]

    private _getFAQData(): Promise<ISPLists> {
      let restQuery = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Frequently Asked Questions')/items?&$select=Id, Title, Answers`;
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
          }
        ]
      };
    }

  //#endregion

}