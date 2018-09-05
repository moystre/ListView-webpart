import { Environment, EnvironmentType, Version } from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneDropdown, PropertyPaneTextField } from '@microsoft/sp-webpart-base';
import * as strings from 'ListViewWebPartStrings';
import { IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { IListViewProps } from './components/IListViewProps';
import ListView from './components/ListView';

export interface IListViewWebPartProps {
  description: string;
  dropdownField: string;
}

export interface ILists {
  value: IList[];
}

export interface IList {
  [key: string]: any;
  Id: string;
  Title: string;
  Description: string;
}

export interface IListItem {
  [key: string]: any;
  listId: string,
  listTitle: string,
  listDescription: string
}

export interface IPropertyLists {
  value: IPropertyList[];
}

export interface IPropertyList {
  key: string,
  text: string
}

export interface IItem {
  [key: string]: any;
  title: string;
}

export default class ListViewWebPart extends BaseClientSideWebPart<IListViewWebPartProps> {
  public listItems: IListItem[];
  public propertyList: IPropertyList[];
  // public isGetItemsFinished: boolean;
  public columns: IColumn[];
  public items: IItem[];

  constructor() {
    super();
    this.columns = [{
      key: 'column1',
      name: 'Title',
      fieldName: 'title',
      minWidth: 200
    }]
  }

  protected async onInit(): Promise<void> {

    this.listItems = await this.getItems();
    console.log(this.listItems);
    console.log('items');
    while (this.listItems == null) {
      /*
       if(!this.isGetItemsFinished) {}
       else {
         break;
       }
       */
    }
    this.propertyList = await this.getPropertyList();
    console.log(this.propertyList);
    console.log('propertyList');

    this.items = [];
  }

  public render(): void {
    const element: React.ReactElement<IListViewProps> = React.createElement(
      ListView,
      {
        description: this.properties.description,
        dropdownField: this.properties.dropdownField,
        columns: this.columns,
        items: this.items
      }
    );
    ReactDom.render(element, this.domElement);
  }

  private async getPropertyList(): Promise<IPropertyList[]> {
    var renderedList: IPropertyList[]
    let i: number = 0;
    var list: {
      key: string,
      text: string
    }[] = [];
    this.listItems.forEach((element: IListItem) => {
      list.push({
        key: i.toString(),
        text: element.listTitle
      })
      i = i + 1;
    });
    renderedList = list;
    return renderedList;
  }

  private async getItems(): Promise<IListItem[]> {
    var renderedList: IListItem[];
    if (Environment.type === EnvironmentType.Local) {
      console.log('Local environment');
      return null;
    }
    else if (Environment.type == EnvironmentType.SharePoint ||
      Environment.type == EnvironmentType.ClassicSharePoint) {
      try {
        let container = null;
        var list: {
          listId: string;
          listTitle: string;
          listDescription: string;
        }[] = [];
        container = await this._getListData();
        container.value.forEach((item: IList) => {
          // console.log(item);
          list.push({
            listId: item.Id,
            listTitle: item.Title,
            listDescription: item.Description
          })
        });
        renderedList = list;
      }
      catch (exception) {
        console.warn(exception);
      }
      return renderedList;
    }
  }

  private _getListData = async (): Promise<ILists> => {
    let returnLists: ILists = null;
    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl
        + `/_api/web/lists?$filter=Hidden eq false`,
        // + `/_api/web/lists/GetByTitle('Collaborators')/items`,
        SPHttpClient.configurations.v1);
      if (!response.ok) {
        throw "Could not fetch list data";
      }
      const lists: ILists = await response.json();
      returnLists = lists;
    } catch (exception) {
      console.warn(exception);
    }
    return returnLists;
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  public componentDidMount(): void {
    //
  }

  public componentDidUpdate(): void {
    console.log(this.properties.dropdownField.toString());
  }

  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    this.listItems = await this.getItems();
    this.propertyList = await this.getPropertyList();
    // this.isGetItemsFinished = true;
  }

  protected onPropertyPaneFieldChanged(): void {
    console.log(this.properties.dropdownField);
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
                  label: 'Description'
                }),
                PropertyPaneDropdown('dropdownField', {
                  label: 'Selected list:',
                  options: this.propertyList
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

// 508 - getEntityFields