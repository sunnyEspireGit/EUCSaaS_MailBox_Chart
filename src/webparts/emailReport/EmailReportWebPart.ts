import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'EmailReportWebPartStrings';
import EmailReport from './components/EmailReport';
import { IEmailReportProps } from './components/IEmailReportProps';
import { sp } from '@pnp/sp';

import {
	IDynamicDataPropertyDefinition,
	IDynamicDataCallables,
} from "@microsoft/sp-dynamic-data";

export interface IEmailReportWebPartProps {
  description: string;
}

export interface IPreferences {
	color?: string;
	date?: Date | null | undefined;
	like?: boolean;
}

export default class EmailReportWebPart extends BaseClientSideWebPart<IEmailReportWebPartProps> implements IDynamicDataCallables{

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public _startDate: string;
  public _endDate: string;
  public _dateRange: string;

  // protected onInit(): Promise<void> {
  //   this._environmentMessage = this._getEnvironmentMessage();

  //   return super.onInit();
  // }

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    this.context.dynamicDataSourceManager.initializeSource(this);

    return super.onInit().then((_) => {
      // other init code may be present

      sp.setup({
        spfxContext: this.context,
      });
    });
  }

  public render(): void {
    const element: React.ReactElement<IEmailReportProps> = React.createElement(
      EmailReport,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        

        onStartDateChanged: this._handleChangeStartDate,
        onEndDateChanged: this._handleChangeEndDate,

        applyDateValue: this._applyDateValue,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _handleChangeStartDate = (date) => {
		this._startDate = date.toISOString();
    console.log("Source startdate",date);
		// notify subscribers that the last name has changed
		this.context.dynamicDataSourceManager.notifyPropertyChanged(
			// Constants.LastNamePropertyId
      "startDate"
		);
	}

  private _handleChangeEndDate = (date) => {
		this._endDate = date.toISOString();
    console.log("Source enddate",date);
		// notify subscribers that the last name has changed
		this.context.dynamicDataSourceManager.notifyPropertyChanged(
			// Constants.LastNamePropertyId
      "endDate"
		);
	}

  public _applyDateValue = (startDate, endDate) =>{

    // this._startDate = startDate;
    // this._endDate = endDate;
    this._dateRange = startDate + " - " + endDate;
    console.log('dateRange....', startDate, endDate);
    this.context.dynamicDataSourceManager.notifyPropertyChanged("dateRange") ;
    // this.context.dynamicDataSourceManager.notifyPropertyChanged("endDate") ;
  }

  /* IDynamicDataCallables implementation*/
	public getPropertyDefinitions(): ReadonlyArray<IDynamicDataPropertyDefinition> {
		return [

      {
				id: "startDate",
				title: "Start Date",
			},
			{
				id: "endDate",
				title: "End Date",
			},
      {
				id: "dateRange",
				title: "Date Range",
			},
			// {
			// 	id: Constants.FirstNamePropertyId,
			// 	title: strings.FirstName,
			// },
			// {
			// 	id: Constants.LastNamePropertyId,
			// 	title: strings.LastName,
			// },
			// {
			// 	id: Constants.PreferencesPropertyId,
			// 	title: strings.Preferences,
			// },
		];
	}

	public getPropertyValue(propertyId: string) {
		switch (propertyId) {
    	case "startDate":
				return this._startDate;
			case "endDate":
				return this._endDate;  

        case "dateRange":
          return (this._dateRange);  
			// case Constants.FirstNamePropertyId:
			// 	return this._firstName;
			// case Constants.LastNamePropertyId:
			// 	return this._lastName;
			// case Constants.PreferencesPropertyId:
			// 	return this._preferences;
		}

		// throw new Error(strings.BadPropertyId);
	}

	/* End of IDynamicDataCallables implementation */

 

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
