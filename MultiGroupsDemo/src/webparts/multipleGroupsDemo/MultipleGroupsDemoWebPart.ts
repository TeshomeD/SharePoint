import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './MultipleGroupsDemoWebPart.module.scss';
import * as strings from 'MultipleGroupsDemoWebPartStrings';

export interface IMultipleGroupsDemoWebPartProps {
  description: string;
  
  productName: string;
  isCertified: boolean;
}

export default class MultipleGroupsDemoWebPart extends BaseClientSideWebPart<IMultipleGroupsDemoWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  protected get disableReactivePropertyChanges():boolean{
    return true;
  }

  public render(): void {
    this.domElement.innerHTML = `
    <div class="${ styles.multipleGroupsDemo }">
      <div class="${styles.container }">
        <div class="${styles.row }">
          <div class="${styles.column }">
            <span class="${styles.title}">Welcome to SharePoint!</span>
            <p class="${styles.subTitle}">Customixe SharePoint experiences using web Parts. </p>
            <p class="${styles.description}"> ${escape(this.properties.description)} </p>

            <p class="${styles.description}"> ${escape(this.properties.productName)} </p>

            <p class="${styles.description}"> ${this.properties.isCertified} </p>
            <a href="https://aka.ms/spfx class="${styles.button}">
              <span class="${styles.label } "> Learn more </span>
            </a>
          </div>
        </div>
      </div>
    </div>
    `;
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
            description: "Page-1"
          },
          groups: [
            {
              groupName: "P1 First Group",
              groupFields: [
                PropertyPaneTextField('productName', {
                  label: "Product Name"
                })
              ]
            },

            {
              groupName: "P1 Second Group",
              groupFields: [
                PropertyPaneToggle('isCertified', {
                  label: "Is Certified?"
                })
              ]
            }
          ],
          displayGroupsAsAccordion:true,
        },
       
        {
          header: {
            description: "Page-2"
          },
          groups: [
            {
              groupName: "P2 First Group",
              groupFields: [
                PropertyPaneTextField('productName', {
                  label: "Product Name"
                })
              ]
            },

            {
              groupName: "P2 Second Group",
              groupFields: [
                PropertyPaneToggle('isCertified', {
                  label: "Is Certified?"
                })
              ]
            }
          ],
          displayGroupsAsAccordion:true,
        },
        {
          header: {
            description: "Page-3"
          },
          groups: [
            {
              groupName: "P3 First Group",
              groupFields: [
                PropertyPaneTextField('productName', {
                  label: "Product Name"
                })
              ]
            },

            {
              groupName: "P3 Second Group",
              groupFields: [
                PropertyPaneToggle('isCertified', {
                  label: "Is Certified?"
                })
              ]
            }
          ],
          displayGroupsAsAccordion:true,
        },
        {
          header: {
            description: "Page-4"
          },
          groups: [
            {
              groupName: "P4 First Group",
              groupFields: [
                PropertyPaneTextField('productName', {
                  label: "Product Name"
                })
              ]
            },

            {
              groupName: "P4 Second Group",
              groupFields: [
                PropertyPaneToggle('isCertified', {
                  label: "Is Certified?"
                })
              ]
            }
          ],
          displayGroupsAsAccordion:true,
        }

      ]//end of pages
    };
  }
}
