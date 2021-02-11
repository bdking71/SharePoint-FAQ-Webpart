import * as React from 'react';
import styles from './FaqSpFxClientSideWebPart.module.scss';
import { IFaqSpFxClientSideWebPartProps } from './IFaqSpFxClientSideWebPartProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class FaqSpFxClientSideWebPart extends React.Component<IFaqSpFxClientSideWebPartProps, {}> {
  public render(): React.ReactElement<IFaqSpFxClientSideWebPartProps> {
    return (
      <div className={ styles.faqSpFxClientSideWebPart }>
        <div className={styles.head}><h1>{escape(this.props.title)}</h1><span>{escape(this.props.description)}</span></div>
        <div className={styles.container}>

        </div>
      </div>
    );
  }
}