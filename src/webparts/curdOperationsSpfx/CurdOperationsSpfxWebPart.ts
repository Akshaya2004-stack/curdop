import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import CurdOperationsSpfx from './components/CurdOperationsSpfx';
import { ICurdOperationsSpfxProps } from './components/ICurdOperationsSpfxProps';

export interface ICurdOperationsSpfxWebPartProps {
  description: string;
  listName: string;
}

export default class CurdOperationsSpfxWebPart extends BaseClientSideWebPart<ICurdOperationsSpfxWebPartProps> {

  private _isDarkTheme: boolean = false;

  public render(): void {
    const element: React.ReactElement<ICurdOperationsSpfxProps> = React.createElement(
      CurdOperationsSpfx,
      {
         description: this.properties.description || 'CRUD Operations With ReactJs',
         listName: this.properties.listName || 'Employees',
        // isDarkTheme: this._isDarkTheme,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }
    this._isDarkTheme = !!currentTheme.isInverted;
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
            description: 'Configure CRUD Web Part'
          },
          groups: [
            {
              groupName: 'Configuration',
              groupFields: [
                PropertyPaneTextField('description', {
                  label: 'Heading Title'
                }),
                PropertyPaneTextField('listName', {
                  label: 'SharePoint List Name',
                  description: 'List containing Employee Name, Hire Date, and Job Description'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
