// Author : Christophe Humbert
// Github : @PathToSharePoint
// Twitter: @Path2SharePoint

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { IPropertyPaneCustomFieldProps, IPropertyPaneField, PropertyPaneFieldType } from '@microsoft/sp-property-pane';

export interface IPropertyPaneWrapBuilderProps {
  component?: any;
  props?: any;
}

export interface IPropertyPaneWrapBuilderInternalProps extends IPropertyPaneWrapBuilderProps, IPropertyPaneCustomFieldProps {
}

export class PropertyPaneWrapBuilder implements IPropertyPaneField<IPropertyPaneWrapBuilderInternalProps> {

  //Properties defined by IPropertyPaneField
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyPaneWrapBuilderInternalProps;

  private elem!: HTMLElement;

  constructor(targetProperty: string, wrapProperties: IPropertyPaneWrapBuilderProps) {
    this.targetProperty = targetProperty;
    this.properties = {
      component: wrapProperties.component,
      props: wrapProperties.props,
      key: targetProperty,
      onRender: this.onRender.bind(this),
      onDispose: this.onDispose.bind(this)
    };
  }

  public render(): void {
    if (!this.elem) {
      return;
    }
    this.onRender(this.elem);
  }

  private onDispose(element: HTMLElement): void {
    ReactDom.unmountComponentAtNode(element);
  }

  private onRender(elem: HTMLElement): void {
    if (!this.elem) {
      this.elem = elem;
    }

    let CustomMGTComponent = this.properties.component;
    let componentProps = this.properties.props;
    ReactDom.render(<CustomMGTComponent {...componentProps}/>, elem);
  }
}

export function PropertyPaneWrap(targetProperty: string, wrapProperties: IPropertyPaneWrapBuilderProps): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  return new PropertyPaneWrapBuilder(targetProperty, wrapProperties);
}