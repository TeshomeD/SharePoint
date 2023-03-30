import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './NewListCreationWpWebPart.module.scss';
import * as strings from 'NewListCreationWpWebPartStrings';

import {SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions} from '@microsoft/sp-http';

export interface INewListCreationWpWebPartProps {
  description: string;
}

export default class NewListCreationWpWebPart extends BaseClientSideWebPart<INewListCreationWpWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    this.domElement.innerHTML = `
    <div class="${styles.newListCreationWp}">
      <h3> Creating a new List Dynamically </h3><br/><br/>
      <p> Please fill out the below details to create a new list programatically </p><br/><br/>

      New List Name:<br/><input type="text" id="txtNewListName" /> <br/><br/>
      New List Discription:<br/><input type="text" id="txtNewListDiscription" /> <br/><br/>

      <input type="button" id="btnCreatNesList" value="Create a New List" /> <br/>

    </div>`;
    this.bindEvents();
  }

  private bindEvents():void {
    this.domElement.querySelector('#btnCreatNesList').addEventListener("click", ()=>{this.createNewList();});
  }
  private createNewList():void {
    var newListName = document.getElementById("txtNewListName")["value"];
    var newListDescription = document.getElementById("txtNewListDiscription")["value"];
    const baseUrl:string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists";
    console.log("baseUrl: ", baseUrl);
    
    const listurl: string= baseUrl + "/GetByTitle('" + newListName + "')";

    this.context.httpClient.get(listurl, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse)=>{
      if(response.status ===200){
        alert('A List already does exist with this name.');
        return;
      }if(response.status===404){
        const listDefinition: any = {
          "Title": newListName,
          "Description": newListDescription,
          "AllowContentTypes": true,
          "BaseTemplate":100,
          "ContentTypesEnabled": true,
        };
         const spHttpClientOptions: ISPHttpClientOptions={
          "body": JSON.stringify(listDefinition)
         };
         this.context.spHttpClient.post(baseUrl, SPHttpClient.configurations.v1,spHttpClientOptions)
         .then((response:SPHttpClientResponse)=>{
            if(response.status===201){
              alert("A new List has been created Successfully.");
            }else{
              alert("Error Message new List has not been created: " +response.status + " - " + response.statusText);
            }
         });
      }else{
              alert("Last Error Message : " +response.status + " - " + response.statusText);
      }
    });

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
