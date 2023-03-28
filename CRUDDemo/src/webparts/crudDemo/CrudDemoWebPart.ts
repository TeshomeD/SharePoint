import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './CrudDemoWebPart.module.scss';
import * as strings from 'CrudDemoWebPartStrings';

import {SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions} from '@microsoft/sp-http';
import { ISoftwareListItem } from './ISoftwareListItem';

export interface ICrudDemoWebPartProps {
  description: string;
}

export default class CrudDemoWebPart extends BaseClientSideWebPart<ICrudDemoWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    this.domElement.innerHTML = `
    <div class="${styles.crudDemo} ">
      <div>
          <table border="5" bgcolor="aqua" >

          <tr>
            <td>Please Enter Software ID</td>
            <td><input type="text" id="txtID"></td>
            <td><input type="submit" id="btnRead" value="Dead Detaile" />
          </tr>

          <tr>
            <td>Software Title</td>
            <td><input type="text" id="txtSoftwarTitle"></td>
          </tr>

          <tr>
            <td>Software Name</td>
            <td><input type="text" id="txtSoftwareName"></td>
          </tr>

          <tr>
            <td>Software Vendor</td>
            <td>
              <select id="ddlSoftwareVendor">
                <option value="Microsoft">Microsoft</option>
                <option value="Sun">Sun</option>
                <option value="Oracle">Oracle</option>
                <option value="Google">Google</option>          
              </select>
          </td>
          </tr>

          <tr>
            <td>Software Version</td>
            <td><input type="text" id="txtSoftwareVersion"></td>
          </tr>

          <tr>
            <td>Software Description</td>
            <td><textarea rows="5" cols="40" id="txtSoftwareDescription"></textarea></td>
          </tr>

          <tr>
            <td colspan="2" align="center">
              <input type="submit" value="Insert Item" id="btnSubmit" />
              <input type="submit" value="Update" id="btnUpdate" />
              <input type="submit" value="Delete" id="btnDelete" />
              <input type="submit" value="Show All Records" id="btnReadAll" />
            </td>
          </tr>                   
        </table>   
      </div>
      <div id="divStatus" />
      
    </div>`;
    this._bindEvents();
  }

  private _bindEvents():void{
    this.domElement.querySelector("#btnSubmit").addEventListener("click",()=>{this.addListItem();});
    this.domElement.querySelector("#btnRead").addEventListener("click", ()=>{this.readListItem();});
  }
  private addListItem():void{
    //we need to creae:
    // get user info
    var softwarTitle = document.getElementById('txtSoftwarTitle')['value'];
    var softwareName  = document.getElementById('txtSoftwareName')['value'];
    var softwareVendor = document.getElementById('ddlSoftwareVendor')['value'];
    var softwareVersion = document.getElementById('txtSoftwareVersion')['value'];
    var softwareDescription = document.getElementById('txtSoftwareDescription')['value'];
    // url
    const stringUrl:string = this.context.pageContext.site.absoluteUrl + "/sites/ItAcademySite/_api/web/lists/getByTitle('SoftwareCatalog')/items";
    // https://spstudent2023.sharepoint.com/sites/ItAcademySite/Lists/SoftwareCatalog/
    // https://spstudent2023.sharepoint.com/_api/web/lists/getByTitle('SoftwareCatalog')/items
    console.log("stringUrl: ", stringUrl);
    const itemBody: any={
      "Title": softwarTitle,
      "SoftwareVendor": softwareVendor,
      "SoftwareDescription": softwareDescription,
      "SoftwareName": softwareName,
      "SoftwareVersion": softwareVersion,
    };
    const spHttpClientOptions: ISPHttpClientOptions = {
      "body": JSON.stringify(itemBody)
    }
    //spHttpRest call
    let statusMessage: Element = this.domElement.querySelector('#divStatus');
    this.context.spHttpClient.post(stringUrl, SPHttpClient.configurations.v1, spHttpClientOptions)
    .then((response: SPHttpClientResponse) => {
      if(response.status ===201) {     
        statusMessage.innerHTML = "List Item has been created successfully.";
        this.clear();
      }else{
         statusMessage.innerHTML = "An error has occured i.e "+response.status + " - "+response.statusText;
      }
    });
    //update the status
  }

  private readListItem():void{
    var softwareId = document.getElementById("txtID")["value"]
    

  }
  private _getListItemById(id: string):Promise<ISoftwareListItem>{
    const readUrl: string = this.context.pageContext.site.absoluteUrl + "/sites/ItAcademySite/_api/web/lists/getbyid('SoftwareCatalog')/items";
    return this.context.spHttpClient.get(readUrl, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) =>{
      return response.json();
    })
  }
  
  private clear(): void{
    document.getElementById('txtSoftwarTitle')['value'] = "";
    document.getElementById('ddlSoftwareVendor')['value'] = "Microsoft";
    document.getElementById('txtSoftwareVersion')['value'] = "";
    document.getElementById('txtSoftwareDescription')['value'] = "";
    document.getElementById('txtSoftwareName')['value'] = "";
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
