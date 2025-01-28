import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDynamicField,
  PropertyPaneDynamicFieldSet,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, IWebPartPropertiesMetadata } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ChartWpWebPartStrings';
import ChartWp from './components/ChartWp';
import { IChartWpProps } from './components/IChartWpProps';
import { sp } from '@pnp/sp';
// import { getSP } from "./pnpjsConfig";
import { DynamicProperty } from '@microsoft/sp-component-base';


export interface IChartWpWebPartProps {
  description: string;
  startDate: DynamicProperty<String>;
  endDate: DynamicProperty<String>;
  dateRange: DynamicProperty<String>;
}

export default class ChartWpWebPart extends BaseClientSideWebPart<IChartWpWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  // protected onInit(): Promise<void> {
  //   this._environmentMessage = this._getEnvironmentMessage();

  //   return super.onInit();
  // }
  
  // protected async onInit(): Promise<void> {
  //   this._environmentMessage = this._getEnvironmentMessage();

  //   super.onInit();

  //   //Initialize our _sp object that we can then use in other packages without having to pass around the context.
  //   //  Check out pnpjsConfig.ts for an example of a project setup file.
  //   getSP(this.context);
  // }

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit().then((_) => {
      // other init code may be present

      sp.setup({
        spfxContext: this.context,
      });
    });
  }


  public render(): void {
    const element: React.ReactElement<IChartWpProps> = React.createElement(
      ChartWp,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        
        startDate:this.properties.startDate,
        endDate: this.properties.endDate,
        dateRange: this.properties.dateRange,
      }
    );

    ReactDom.render(element, this.domElement);
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

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  // protected get propertiesMetadata():IWebPartPropertiesMetadata {
  //   return{
  //     startDate:{dynamicPropertyType:"string"}
  //   };
  // }

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
                  label: strings.DescriptionFieldLabel,
                }),
                // PropertyPaneDynamicFieldSet({label: "Select Source Webpart", fields:[PropertyPaneDynamicField("startDate", {label: "startDate",})]}),
                // PropertyPaneDynamicFieldSet({label: "Select Source Webpart", fields:[PropertyPaneDynamicField("endDate", {label: "endDate",})]}),

                PropertyPaneDynamicField("startDate", {
									label: "Start Date",
								}),
                PropertyPaneDynamicField("endDate", {
									label: "End Date",
								}),
                PropertyPaneDynamicField("dateRange", {
                  label: "Date Changes",
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
