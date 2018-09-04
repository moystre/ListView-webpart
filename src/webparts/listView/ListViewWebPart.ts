import { Environment, EnvironmentType, Version } from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneDropdown, PropertyPaneTextField } from '@microsoft/sp-webpart-base';
import * as strings from 'ListViewWebPartStrings';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { IListViewProps } from './components/IListViewProps';
import ListView from './components/ListView';

export interface IListViewWebPartProps {
  description: string;
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

export default class ListViewWebPart extends BaseClientSideWebPart<IListViewWebPartProps> {
  public items: IListItem[];
  public propertyList: IPropertyList[];
  // public isGetItemsFinished: boolean;

  constructor() {
    super();
  }

  protected async onInit(): Promise<void> {

    this.items = await this.getItems();
    console.log(this.items);
    console.log('items');
    while (this.items == null) {
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
  }

  public render(): void {
    const element: React.ReactElement<IListViewProps> = React.createElement(
      ListView,
      {
        description: this.properties.description
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
    this.items.forEach((element: IListItem) => {
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

  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    this.items = await this.getItems();
    this.propertyList = await this.getPropertyList();
    // this.isGetItemsFinished = true;
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
                PropertyPaneDropdown('test', {
                  label: 'Dropdown',
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


