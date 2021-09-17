import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'TeamsAuthIssueReproWpWebPartStrings';
import TeamsAuthIssueReproWp from './components/TeamsAuthIssueReproWp';
import { ITeamsAuthIssueReproWpProps } from './components/ITeamsAuthIssueReproWpProps';

export interface ITeamsAuthIssueReproWpWebPartProps {
  description: string;
}

export default class TeamsAuthIssueReproWpWebPart extends BaseClientSideWebPart<ITeamsAuthIssueReproWpWebPartProps> {

  protected tokenResponse: string;

  public render(): void {
    const element: React.ReactElement<ITeamsAuthIssueReproWpProps> = React.createElement(
      TeamsAuthIssueReproWp,
      {
        description: this.properties.description,
        context: this.context
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
