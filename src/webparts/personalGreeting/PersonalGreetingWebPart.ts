import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,  
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'PersonalGreetingWebPartStrings';
import PersonalGreeting from './components/PersonalGreeting';
import { IPersonalGreetingProps } from './components/IPersonalGreetingProps';

import { getSP, CustomListener } from './pnpjsConfig';
import { Logger, LogLevel } from "@pnp/logging";
import UserProfileService from './services/UserProfileService';

export interface IPersonalGreetingWebPartProps {
  message: string;
  optionalMessage: string;  
  primaryTextColor: string;
  secondaryTextColor: string;
  tertiaryTextColor: string;
  primaryTextSize: number;
  secondaryTextSize: number;
  tertiaryTextSize: number;
  position: string;
  lineHeight: string;  
}

export default class PersonalGreetingWebPart extends BaseClientSideWebPart<IPersonalGreetingWebPartProps> {  
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  private _LOG_SOURCE: string = "PersonalGreetingWebPart";
  private _userProfileService: UserProfileService;  
  private _preferredName: string;
  private _todayDate: string = '';    
  
  protected async onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    await super.onInit();

    Logger.activeLogLevel = LogLevel.Warning;
    Logger.subscribe(new CustomListener());
    
    // Initialize our _sp object that we can then use in other packages without having to pass around the context.    
    getSP(this.context);

    await this._getPreferredName();
    this._getCurrentDate();
  }
  
  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }
    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }
  
  private _getPreferredName = async (): Promise<void> => {
    try {
      if(!this._userProfileService) {
        this._userProfileService = new UserProfileService();
      }
      
      await this._userProfileService.retrieveCurrentUserProfiles();
      this._preferredName= await this._userProfileService.getCurrentUserPreferredName();      
      // console.log("_preferredName: "+this._preferredName);
        
    } catch (err) {
      Logger.write (`${this._LOG_SOURCE} (_loadCurrentUserGroups) - ${LogLevel.Error}\n${err.message}`, LogLevel.Error);      
    }
  }

  private _getCurrentDate(): void {    
    if(window.location.href.toString().indexOf('fr') > -1) {
      const today: string = new Date().toLocaleDateString('fr-FR', {
        weekday: "short",
        day : 'numeric',
        month : 'short',
        year : 'numeric'
      });      
      this._todayDate = today;
    } else {
      const arrToday: string[]  = new Date().toDateString().split(' ');// Fri Jul 29 2022
      const todayFormatted: string = arrToday[0] + ', ' + arrToday[1] + ' ' + arrToday[2] + ', ' + arrToday[3];      
      this._todayDate = todayFormatted;
    }
    // console.log(this._todayDate);
  }
  
  public render(): void {
    const element: React.ReactElement<IPersonalGreetingProps> = React.createElement(
      PersonalGreeting,
      {
        currentDate: this._todayDate,
        message: this.properties.message,
        optionalMessage: this.properties.optionalMessage,
        primaryTextColor: this.properties.primaryTextColor,
        secondaryTextColor: this.properties.secondaryTextColor,
        tertiaryTextColor: this.properties.tertiaryTextColor,
        primaryTextSize: this.properties.primaryTextSize,
        secondaryTextSize: this.properties.secondaryTextSize,
        tertiaryTextSize: this.properties.tertiaryTextSize,
        position: this.properties.position,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this._preferredName
      }
    );
    
    ReactDom.render(element, this.domElement);

    //Dynamically apply css style to the variable of .customClass
    this.domElement.style.setProperty('--lineHeight', this.properties.lineHeight || null);
    // console.log("*** Applied LineHeight " + this.properties.lineHeight.toString());
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
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
                PropertyPaneTextField('message', {
                  label: "Greeting Message"              
                }),
                PropertyFieldColorPicker('primaryTextColor', {
                  label: 'Greeting Message Color',
                  properties: this.properties,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  selectedColor: this.properties.primaryTextColor,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'primaryTextColor'
                }),
                PropertyPaneSlider('primaryTextSize', {
                  label: 'Greeting Message Size',
                  min: 1,
                  max: 30,
                  step: 1,
                  value: 20
                }),
                PropertyPaneTextField('optionalMessage', {
                  label: "Optional Message",               
                  multiline: true
                }),
                PropertyFieldColorPicker('secondaryTextColor', {
                  label: 'Optional Message Color',
                  properties: this.properties,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  selectedColor: this.properties.secondaryTextColor,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'secondaryTextColor'
                }),
                PropertyPaneSlider('secondaryTextSize', {
                  label: 'Optional Message Size',
                  min: 0,
                  max: 30,
                  step: 1,
                  value: 10
                }),
                PropertyFieldColorPicker('tertiaryTextColor', {
                  label: 'Date Text Color',
                  properties: this.properties,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  selectedColor: this.properties.tertiaryTextColor,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'tertiaryTextColor'
                }),
                PropertyPaneSlider('tertiaryTextSize', {
                  label: 'Date Text Size',
                  min: 0,
                  max: 30,
                  step: 1,
                  value: 14
                }),
                PropertyPaneDropdown('position', {
                  label: 'Text Position',
                  selectedKey: 'left',
                  options: [{
                    key: 'left',
                    text: 'left'
                  },
                  {
                    key: 'center',
                    text: 'center'
                  },
                  {
                    key: 'right',
                    text: 'right'
                  }
                  ]
                }),
                PropertyPaneSlider('lineHeight', {
                  label: 'Line Height',
                  min: 0,
                  max: 3,
                  step: 0.1,
                  value: 0.2
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}