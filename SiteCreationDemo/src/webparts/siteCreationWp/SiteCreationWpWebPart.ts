import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SiteCreationWpWebPart.module.scss';
import * as strings from 'SiteCreationWpWebPartStrings';

import {SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions} from '@microsoft/sp-http';

export interface ISiteCreationWpWebPartProps {
  description: string;
}

export default class SiteCreationWpWebPart extends BaseClientSideWebPart<ISiteCreationWpWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    this.domElement.innerHTML = `
    <div class="${styles.siteCreationWp }">
      <h1>Create a New Subsite </h1>
      <p>Please fill the below details to create a new subsite. </p><br/>
      Sub Site Title: </br><input type= text id='textSubSiteTitle' /><br/>
      Sub Site URL: </br><input type= text id='textSubSiteUrl' /><br/>
      Sub Site Description: </br><textarea id='textSubSiteDescription' rows="5" cols="30"></textarea></br>
      <br/>
      <input type="button" id="btnCreateSubSite" value="Create Sub Site" /><br/>
      
    </div>`;

    this.bindEvents();
  }
  private bindEvents():void {
    this.domElement.querySelector('#btnCreateSubSite').addEventListener('click',() => {this.createSubSite();});
  }
  private createSubSite():void{
    let subSiteTitle = document.getElementById("textSubSiteTitle")['value'];
    let subSiteUrl= document.getElementById("textSubSiteUrl")['value']; 
    let subSiteDescription = document.getElementById("textSubSiteDescription")['value'];

    const url: string = this.context.pageContext.web.absoluteUrl + "/_api/web/webinfos/add"; 
    console.log("url: ",url);

    const spHttpClientOptions: ISPHttpClientOptions={
        body: `{
          "parameters":{
            "@odata.type": "SP.WebInfoCreationInformation",
            "Title": "${subSiteTitle}",
            "Url": "${subSiteUrl}",
            "Description": "${subSiteDescription}",
            "Language": 1033,
            "WebTemplate": "STS#0",
            "UseUniquePermissions": true
              }
            }`
    };
    
    this.context.spHttpClient.post(url, SPHttpClient.configurations.v1,spHttpClientOptions)
    .then((response: SPHttpClientResponse)=>{
       if(response.status === 200){
        alert("New Subsite has been created successfully");
       }else{
        alert("Error Message: " +response.status + " - "+ response.statusText);
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
