import { escape } from '@microsoft/sp-lodash-subset';
import * as React from 'react';
import { IListViewProps } from './IListViewProps';
import styles from './ListView.module.scss';

export default class ListView extends React.Component<IListViewProps, {}> {
  public render(): React.ReactElement<IListViewProps> {
    return (
      <div className={styles.listView}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to SharePoint!</span>
              <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>
              <p className={styles.description}>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={styles.button}>
                <span className={styles.label}>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
