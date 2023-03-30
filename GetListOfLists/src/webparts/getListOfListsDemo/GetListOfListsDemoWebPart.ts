import { Environment, EnvironmentType, Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './GetListOfListsDemoWebPart.module.scss';
import * as strings from 'GetListOfListsDemoWebPartStrings';

import {SPHttpClient, SPHttpClientResponse} from '@microsoft/sp-http';

export interface IGetListOfListsDemoWebPartProps {
  description: string;
}
export interface ISharePointList{
  Title: string;
  Id: string;
}
export interface ISharePointLists{
  value: ISharePointList[];
}


export default class GetListOfListsDemoWebPart extends BaseClientSideWebPart<IGetListOfListsDemoWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  //private html: string='';

  private _getListOfLists(): Promise<ISharePointLists>{
    const url: string = this.context.pageContext.web.absoluteUrl +"/_api/web/lists";
    console.log("url: ", url);
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
            //this.context.pageContext.web.absoluteUrl +`/_api/web/lists?$filter=Hidden eq false`, SPHttpClient.configurations.v1)
            
      .then((response: SPHttpClientResponse)=>{
        return response.json();
      });
  }

  private _getAndRenderLists(): void{
    // if it is local do nothing
    if(Environment.type === EnvironmentType.Local){}
    else if(Environment.type == EnvironmentType.ClassicSharePoint ||
            Environment.type == EnvironmentType.SharePoint){
      this._getListOfLists().then((response)=>{
        this._renderListOfLists(response.value);
      });
    }
  }
  private _renderListOfLists(items: ISharePointList[]):void {
    let html: string=``;
    items.forEach((item: ISharePointList)=>{
      html+=`
        <ul className="${styles.list}">
          <li className="${styles.listItem}"> <span class="ms-font-l"> ${item.Title} </span> </li>
          <li className="${styles.listItem}"> <span class="ms-font-l"> ${item.Id} </span> </li>
        </ul>;
      `;
    });
    //console.log("html: ",this.html);
    const listPlaceholder: Element = this.domElement.querySelector('#SPListPlaceHolder');
    console.log("listPlaceholder: ", listPlaceholder)
    listPlaceholder.innerHTML=html;
  }

 

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    this._getAndRenderLists();
    this.domElement.innerHTML = `
      <div class="${ styles.getListOfListsDemo }">
    <div class="${ styles.container }">
      <div class="${ styles.row }">
        <div class="${ styles.column }">
          <span class="${ styles.title }">Welcome to SharePoint!</span>
  <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
    <p class="${ styles.description }">${escape(this.properties.description)}</p>
      <a href="https://aka.ms/spfx" class="${ styles.button }">
        <span class="${ styles.label }">Learn more</span>
          </a>
          </div>
          </div>
          </div>

          <div id="SPListPlaceHolder">

          </div>
  
          </div>`;
          

  }


  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

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
}
