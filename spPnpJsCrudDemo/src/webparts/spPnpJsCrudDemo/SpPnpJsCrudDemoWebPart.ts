import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SpPnpJsCrudDemoWebPart.module.scss';
import * as strings from 'SpPnpJsCrudDemoWebPartStrings';

import * as pnp from 'sp-pnp-js';

export interface ISpPnpJsCrudDemoWebPartProps {
  description: string;
}

export default class SpPnpJsCrudDemoWebPart extends BaseClientSideWebPart<ISpPnpJsCrudDemoWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit().then(_ =>{
      pnp.setup({
        spfxContext:this.context
      });
    });
  }

  public render(): void {
    this.domElement.innerHTML = `
    <div class="${styles.spPnpJsCrudDemo} ">
      <div>
          <table border="5" bgcolor="aqua" >

          <tr>
            <td>Please Enter Software ID</td>
            <td><input type="text" id="txtID"></td>
            <td><input type="submit" id="btnRead" value="Read Detaile" />
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
      <h2>Get All List Items </h2>
      <hr/>
      <div id="spListData" />
    </div>`;
    this._bindEvents();
  }

  private _bindEvents():void{
    this.domElement.querySelector('#btnSubmit').addEventListener('click', () => {this.addListItem();});
    this.domElement.querySelector('#btnRead').addEventListener('click', () => {this.readListItem();});
    this.domElement.querySelector('#btnUpdate').addEventListener('click', () =>{this.updateLIstItem();});
    this.domElement.querySelector('#btnDelete').addEventListener('click', () =>{this.deleteLIstItem();});
    this.domElement.querySelector('#btnReadAll').addEventListener('click', () =>{this.readAllItems();});
  }
  private deleteLIstItem():void{
    const id = document.getElementById("txtID")['value'];
    pnp.sp.web.lists.getByTitle('SoftwareCatalogFromCommunicationSite').items.getById(id).delete()
    .then(r=>{
      alert("Item deleted successfuly");
    }).catch(error=>{
      this.domElement.querySelector('#divStatus').innerHTML="Something went wrong";
    })
  }
  private readAllItems(): void{
    let html: string='<table border=1 width=100% style="border-collapse: collapse;">';
      html += '<th> Title </th> <th> Vendor </th> <th> Description </th> <th> Name </th> <th> Version </th>';
      pnp.sp.web.lists.getByTitle('SoftwareCatalogFromCommunicationSite').items.get()
      .then((items:any[])=>{
        items.forEach(function(item){
          html += `<tr>
            <td>${item["Title"]}</td>
            <td>${item["SoftwareVendor"]}</td>
            <td>${item["SoftwareDescription"]}</td>
            <td>${item["SoftwareName"]}</td>
            <td>${item["SoftwareVersion"]}</td>
          </tr> `
        });
        html += '</table>';
        const listContainer: Element = this.domElement.querySelector('#spListData');
        listContainer.innerHTML = html;
      })
  }
  private updateLIstItem():void{
    var title = document.getElementById('txtSoftwarTitle')['value'];
    var softwareName = document.getElementById('txtSoftwareName')['value'];
    var softwareVendor = document.getElementById('ddlSoftwareVendor')['value'];
    var softwareVersion = document.getElementById('txtSoftwareVersion')['value'];
    var softwareDescription = document.getElementById('txtSoftwareDescription')['value'];
    
    const id = document.getElementById("txtID")['value'];
    pnp.sp.web.lists.getByTitle('SoftwareCatalogFromCommunicationSite').items.getById(id).update({
      Title:title,
      SoftwareName: softwareName,
      SoftwareVendor: softwareVendor,
      SoftwareVersion: softwareVersion,
      SoftwareDescription:softwareDescription
    }).then(r=>{
      alert("Details updated");
    })
    

  }
  private readListItem():void{
      const id = document.getElementById("txtID")['value'];
      pnp.sp.web.lists.getByTitle('SoftwareCatalogFromCommunicationSite').items.getById(id).get()
      .then((item: any)=>{
        document.getElementById('txtSoftwarTitle')['value'] = item["Title"];
        document.getElementById('txtSoftwareName')['value'] = item["SoftwareName"];
        document.getElementById('ddlSoftwareVendor')['value'] = item["SoftwareVendor"];
        document.getElementById('txtSoftwareVersion')['value'] = item["SoftwareVersion"];
        document.getElementById('txtSoftwareDescription')['value'] = item["SoftwareDescription"];
      }).catch(error=>{
        let message: Element= this.domElement.querySelector('#divStatus');
        message.innerHTML="Unable to fetch data " + error.status +" - "+error.message
      });
  }
  private addListItem():void{
    var softwarTitle = document.getElementById('txtSoftwarTitle')['value'];
    var softwareName  = document.getElementById('txtSoftwareName')['value'];
    var softwareVendor = document.getElementById('ddlSoftwareVendor')['value'];
    var softwareVersion = document.getElementById('txtSoftwareVersion')['value'];
    var softwareDescription = document.getElementById('txtSoftwareDescription')['value'];

    const siteurl: string = this.context.pageContext.site.absoluteUrl + "/_api/web/lists/getByTitle('SoftwareCatalog')/items";
    const _baseUrl: string = this.context.pageContext.site.absoluteUrl + "/sites/ItAcademySite/_api/web/lists/getByTitle('SoftwareCatalog')/items";
    console.log(" pnp.sp.web.lists: ", pnp.sp.web.lists);  
    console.log(" this.context.pageContext.site.absoluteUrl: ", this.context.pageContext.site.absoluteUrl);

    pnp.sp.web.lists.getByTitle('SoftwareCatalogFromCommunicationSite').items.add({
      Title: softwarTitle,
      SoftwareVendor: softwareVendor,
      SoftwareDescription: softwareDescription,
      SoftwareName: softwareName,
      SoftwareVersion: softwareVersion,
    }).then(r=>{
      alert("Success");
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
