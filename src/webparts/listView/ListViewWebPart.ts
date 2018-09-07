import { Environment, EnvironmentType, Version } from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneDropdown } from '@microsoft/sp-webpart-base';
import * as strings from 'ListViewWebPartStrings';
import { IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { IListViewProps } from './components/IListViewProps';
import ListView from './components/ListView';

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Id: number;
  Title: string;
  Created: Date;
  wpSite: string;
  wpDescription: string;
  wpBusinessModule: string;
}

export interface IListViewWebPartProps {
  description: string;
  dropdownField: string;
}

export interface IListsFromSite {
  value: IListFromSiteAsItem[];
}

export interface IListFromSiteAsItem {
  [key: string]: any;
  Id: string;
  Title: string;
  Description: string;
}

export interface IRenderedListsFromSite {
  [key: string]: any;
  listId: string;
  listTitle: string;
  listDescription: string;
}

export interface IDropDownLists {
  value: IDropDownList[];
}

export interface IDropDownList {
  key: string;
  text: string;
}

export interface IItem {
  [key: string]: any;
  Id: number;
  title: string;
  created: Date;
  site: string;
  description: string;
  businessModule: string;
}

export interface IColumn {
  key: string;
  name: string;
  fieldName: string;
  minWidth: number;
  isResizable: boolean;
  calculatedWidth: any;
}

export default class ListViewWebPart extends BaseClientSideWebPart<IListViewWebPartProps> {
  public renListsFromSite: IRenderedListsFromSite[];
  public dropDownList: IDropDownList[];
  public dropDownfieldName: string = '';
  public columns: IColumn[];
  public items: IItem[];

  constructor() {
    super();
    this.columns = [{
      key: 'column0',
      name: 'Id',
      fieldName: 'Id',
      minWidth: 30,
      maxWidth: 30,
      isResizable: true
    },
    {
      key: 'column1',
      name: 'Title',
      fieldName: 'title',
      minWidth: 150,
      maxWidth: 200,
      isResizable: true
    },
    {
      key: 'column2',
      name: 'Created',
      fieldName: 'created',
      minWidth: 150,
      maxWidth: 150,
      isResizable: true
    },
    {
      key: 'column3',
      name: 'Site',
      fieldName: 'site',
      minWidth: 100,
      maxWidth: 100,
      isResizable: true
    }]
  }

  protected async onInit(): Promise<void> {
    this.renListsFromSite = await this.getRenderedListOfLists();
    while (this.renListsFromSite == null) {
      /* if(!this.isGetItemsFinished) {} */
    }
    this.dropDownList = await this.getSelectionList();
    this.items = [];
    this.items = await this.getItems();
    this.dropDownfieldName = this.getListNameByKey(this.properties.dropdownField);
    console.log(this.dropDownfieldName);
    await this.refreshItems();
  }

  public render(): void {
    const element: React.ReactElement<IListViewProps> = React.createElement(
      ListView,
      {
        description: this.properties.description,
        dropdownField: this.dropDownfieldName,
        columns: this.columns,
        items: this.items
      }
    );
    ReactDom.render(element, this.domElement);
  }

  private async refreshItems(): Promise<IItem[]> {
    return this.items = await this.getItems();
  }

  private async getSelectionList(): Promise<IDropDownList[]> {
    var selectionList: IDropDownList[]
    let i: number = 0;
    var list: {
      key: string,
      text: string
    }[] = [];
    this.renListsFromSite.forEach((element: IRenderedListsFromSite) => {
      list.push({
        key: i.toString(),
        text: element.listTitle
      })
      i = i + 1;
    });
    selectionList = list;
    return selectionList;
  }

  private async getRenderedListOfLists(): Promise<IRenderedListsFromSite[]> {
    var renderedListsFromSite: IRenderedListsFromSite[];
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
        container = await this._getListsFromSite();
        container.value.forEach((item: IListFromSiteAsItem) => {
          //console.log(item);
          list.push({
            listId: item.Id,
            listTitle: item.Title,
            listDescription: item.Description
          })
        });
        renderedListsFromSite = list;
      }
      catch (exception) {
        console.warn(exception);
      }
      return renderedListsFromSite;
    }
  }

  private _getListsFromSite = async (): Promise<IListsFromSite> => {
    let listsFromSite: IListsFromSite = null;
    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl
        + `/_api/web/lists?$filter=Hidden eq false`,
        // + `/_api/web/lists/GetByTitle('Collaborators')/items`,
        SPHttpClient.configurations.v1);
      if (!response.ok) {
        throw "Could not fetch list data";
      }
      const lists: IListsFromSite = await response.json();
      listsFromSite = lists;
    } catch (exception) {
      console.warn(exception);
    }
    return listsFromSite;
  }

  public getListNameByKey(key: string): string {
    let _key: string = key;
    let _text: string = '';
    this.dropDownList.forEach(element => {
      if (element.key == _key) {
        _text = element.text;
      }
    });
    return _text;
  }

  private async getItems(): Promise<IItem[]> {
    var renderedList: IItem[];
    if (Environment.type === EnvironmentType.Local) {
      console.log('Local environment');
      return null;
    }
    else if (Environment.type == EnvironmentType.SharePoint ||
      Environment.type == EnvironmentType.ClassicSharePoint) {
      try {
        let container = null;
        var list: {
          Id: number,
          title: string,
          created: Date
          site: string,
          description: string,
          businessModule: string
        }[] = [];
        if(this.properties.dropdownField == null) {
          this.properties.dropdownField = '1';
        } else {
          container = await this._getListData(this.getListNameByKey(this.properties.dropdownField));
        }
        container.value.forEach((item: ISPList) => {
          console.log(item);
          list.push({
            Id: item.Id,
            title: item.Title,
            created: item.Created,
            site: item.wpSite,
            description: item.wpDescription,
            businessModule: item.wpBusinessModule
          })
        });
        renderedList = list;
        console.log(list);
      }
      catch (exception) {
        console.warn(exception);
      }
      return renderedList;
    }
  }

  private _getListData = async (listName: string): Promise<ISPLists> => {
    let returnLists: ISPLists = null;
    let _listName: string = listName;
    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl
        /*  + `/_api/web/lists?$filter=Hidden eq false`, */
        + `/_api/web/lists/GetByTitle('` + _listName + `')/items`,
        SPHttpClient.configurations.v1);
      if (!response.ok) {
        throw "Could not fetch list data";
      }
      const lists: ISPLists = await response.json();
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
  }

  public componentDidUpdate(): void {
  }

  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    this.refreshItems();
    this.dropDownList = await this.getSelectionList();
  }

  protected async onPropertyPaneFieldChanged(): Promise<void> {
    await this.refreshItems();
    this.dropDownfieldName = this.dropDownList[this.properties.dropdownField].text.toString();
    console.log(this.dropDownfieldName);
    this.render();
    return null;
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
                PropertyPaneDropdown('dropdownField', {
                  label: 'Selected list:',
                  options: this.dropDownList
                })
              ]
            }
          ]
        }
      ]
    };
  }
}