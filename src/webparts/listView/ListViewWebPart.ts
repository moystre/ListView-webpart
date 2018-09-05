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
  listId: string,
  listTitle: string,
  listDescription: string
}

export interface IDropDownLists {
  value: IDropDownList[];
}

export interface IDropDownList {
  key: string,
  text: string
}

export interface IItem {
  [key: string]: any;
  title: string;
}

export default class ListViewWebPart extends BaseClientSideWebPart<IListViewWebPartProps> {
  public renListsFromSite: IRenderedListsFromSite[];
  public dropDownList: IDropDownList[];
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
    this.renListsFromSite = await this.getRenderedListOfLists();
    while (this.renListsFromSite == null) {
      /* if(!this.isGetItemsFinished) {} */
    }
    this.dropDownList = await this.getSelectionList();
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
          console.log(item);
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
    // this.listItems = await this.getRenderedListOfLists();
    this.dropDownList = await this.getSelectionList();
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

// 508 - getEntityFields