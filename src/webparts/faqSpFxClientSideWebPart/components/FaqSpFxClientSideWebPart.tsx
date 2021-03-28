//#region [imports]

  import * as React from 'react';
  import { useState } from 'react';
  import * as ReactDom from 'react-dom';
  import styles from './FaqSpFxClientSideWebPart.module.scss';
  import { IFaqSpFxClientSideWebPartProps } from './IFaqSpFxClientSideWebPartProps';
  import { escape } from '@microsoft/sp-lodash-subset';
  import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
  import {SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration} from '@microsoft/sp-http';
  import { SPComponentLoader } from '@microsoft/sp-loader';
  import ReactHtmlParser from 'react-html-parser'; 

  import * as jQuery from 'jquery';
  import 'jqueryui';
//#endregion

//#region [exports]

interface iState {
  data: any;
}

  export interface iSPFaqListdata {
    value: iSPFaqListItem[];
  }

  export interface iSPFaqListItem {
    Id: number;
    Title: string; 
    Answers: string;
  }

//#endregion 

//#region [constants]

  const accordionOptions: JQueryUI.AccordionOptions = {
    animate: true,
    collapsible: false,
    icons: {
      header: 'ui-icon-circle-arrow-e',
      activeHeader: 'ui-icon-circle-arrow-s'
    }
  };

  const Accordion = ({children}) => { 
    return (
      <div className="accordion">
        {children}
      </div>
    ); 
  };

  const FaqItem = ({question, answer}) => {
    return ( 
      <React.Fragment>
        <h3>{escape(question)}</h3>
        <div><p>{ReactHtmlParser (answer)}</p></div>
      </React.Fragment>
    );
  };

  const WebpartHeader = ({title, description}) => {
    return (
      <div className={styles.head}>
        <h1>{escape(title)}</h1>
        <span>{escape(description)}</span>
      </div>
    );
  };

//#endregion

export default class FaqSpFxClientSideWebPart extends 
  React.Component<IFaqSpFxClientSideWebPartProps, iState> {
  
  constructor(props:IFaqSpFxClientSideWebPartProps) {
      super(props);
      this.state={
        data:[]
      };
      //SPComponentLoader.loadCss('https://code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css');
      SPComponentLoader.loadCss('https://code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css');
    }

  public componentDidUpdate() {
    jQuery('.accordion').accordion(accordionOptions); 
  }

  public componentDidMount() {
    let accordionDiv: HTMLElement = document.getElementById("accordion");
    this.getData().then((response) => {
      this.setState({data: response.value});
    });   
   }

  public render(): React.ReactElement<IFaqSpFxClientSideWebPartProps> { 
    let myData = this.state.data;  
    return (
      <div className={ styles.faqSpFxClientSideWebPart }>
        <WebpartHeader title={this.props.title} description={this.props.description} />
        <Accordion>{myData.map((item) => {
          return <FaqItem question={item.Title} answer={item.Answers}></FaqItem>; 
        })}
        </Accordion> 
      </div>
    );   
  }

  private getData(): Promise<iSPFaqListdata> {
    let restQuery = `${this.props.siteURL}/_api/web/lists/getbytitle('Frequently Asked Questions')/items?&$select=Id, Title, Answers`;
    return this.props.x.context.spHttpClient.get(restQuery ,SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });   
  }

}