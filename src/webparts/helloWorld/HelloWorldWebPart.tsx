import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneHorizontalRule,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'HelloWorldWebPartStrings';
import HelloWorld from './components/HelloWorld';
import { IHelloWorldProps } from './components/IHelloWorldProps';
import { CustomPropertyPane } from './components/CustomPropertyPane';
import { update } from '@microsoft/sp-lodash-subset';
import { PropertyPaneHost } from 'property-pane-portal';

import { Providers, SharePointProvider } from '@microsoft/mgt';

export interface IHelloWorldWebPartProps {
  description: string;
  mgtPerson: string;
  mgtPeoplePicker: string;
  mgtGroupPicker: string;
  mgtTeamsChannelPicker: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit().then(_ => {
  
      // other init code may be present
      Providers.globalProvider = new SharePointProvider(this.context);
      
    });
  }

  public render(): void {

    const wpProps = {
      properties: this.properties,
      isDarkTheme: this._isDarkTheme,
      environmentMessage: this._environmentMessage,
      hasTeamsContext: !!this.context.sdks.microsoftTeams,
      userDisplayName: this.context.pageContext.user.displayName
    };

    const customPropertyPaneProperties = {
      context: this.context,
      properties: this.properties,
      updateWebPartProperty: this.updateWebPartProperty.bind(this),
    };

    ReactDom.render(
      <>
        {/* Web Part content */}
        <HelloWorld {...wpProps} />
        {/* Property Pane custom controls */}
        <CustomPropertyPane {...customPropertyPaneProperties} />
      </>,
      this.domElement);
  }

  public updateWebPartProperty(property, value) {

    update(this.properties, property, () => value);
    this.render();
  
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

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneConfigurationStart() {
  }
  
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

    const hostProperties = {
      context: this.context
    };

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
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneHost('mgtPerson', hostProperties),
                PropertyPaneHorizontalRule(),
                PropertyPaneHost('mgtPeoplePicker', hostProperties),
                PropertyPaneHorizontalRule(),
                PropertyPaneHost('mgtGroupPicker', hostProperties),
                PropertyPaneHorizontalRule(),
                PropertyPaneHost('mgtTeamsChannelPicker', hostProperties),
                PropertyPaneHorizontalRule()
              ]
            }
          ]
        }
      ]
    };
  }
}