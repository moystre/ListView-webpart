import { escape } from '@microsoft/sp-lodash-subset';
import { CheckboxVisibility, DetailsList } from 'office-ui-fabric-react/lib/DetailsList';
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
              <p className={styles.description}>{escape(this.props.description)}</p>
              <p className="ms-font-l ms-fontColor-white">Dropdown selected value: {this.props.dropdownField}</p>
            </div>
          </div>
        </div>
        <div>
          <hr></hr>
          <DetailsList
            items={this.props.items}
            columns={this.props.columns}
            checkboxVisibility={CheckboxVisibility.onHover}
            compact={true}>
          </DetailsList>
        </div>
      </div>

    );
  }
}
