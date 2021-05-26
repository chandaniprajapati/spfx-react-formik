import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ReactFormikWebPartStrings';
import ReactFormik from './components/ReactFormik';
import { IReactFormikProps } from './components/IReactFormikProps';
import { sp } from "@pnp/sp/presets/all";

export interface IReactFormikWebPartProps {
  description: string;
  siteUrl: string;
}

export default class ReactFormikWebPart extends BaseClientSideWebPart<IReactFormikWebPartProps> {

  protected onInit(): Promise<void> {

    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
    });
  }

  public render(): void {
    const element: React.ReactElement<IReactFormikProps> = React.createElement(
      ReactFormik,
      {
        description: this.properties.description,
        context: this.context,
        siteUrl: this.properties.siteUrl ? this.properties.siteUrl : this.context.pageContext.web.absoluteUrl
      }
    );

    ReactDom.render(element, this.domElement);
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
