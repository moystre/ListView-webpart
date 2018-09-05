import { CheckboxVisibility, DetailsList } from 'office-ui-fabric-react/lib/DetailsList';
import * as React from 'react';
import { IListViewProps } from './IListViewProps';
import styles from './ListView.module.scss';

export default class ListView extends React.Component<IListViewProps, {}> {
  public render(): React.ReactElement<IListViewProps> {
    return (
      <div className={styles.listView}>
        <div>
        <span className={styles.title}>{this.props.dropdownField}</span>
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
