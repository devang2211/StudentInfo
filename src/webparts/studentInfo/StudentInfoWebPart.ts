import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'StudentInfoWebPartStrings';
import StudentInfo from './components/StudentInfo';
import { IStudentInfoProps } from './components/IStudentInfoProps';

export interface IStudentInfoWebPartProps {
  description: string;
  // description: string;
  // successMessage: string;
  // requiredFieldMessage: string;
  // errorMessage: string;
  // apsNumberExiteMessage: string;
  // readUserPermissionMessage: string;
  // trainingTitle: string;
  // trainingDocumentTitle: string;  
}

export default class StudentInfoWebPart extends BaseClientSideWebPart<IStudentInfoWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IStudentInfoProps> = React.createElement(
      StudentInfo,
      {
        description: this.properties.description,
        // successMessage: this.properties.successMessage,
        siteAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
        // requiredFieldMessage: this.properties.requiredFieldMessage,
        // errorMessage: this.properties.errorMessage,
        // apsNumberExiteMessage: this.properties.apsNumberExiteMessage,
        context: this.context,
        userPermissions: this.context.pageContext.web.permissions,
        // readUserPermissionMessage: this.properties.readUserPermissionMessage,
        // trainingTitle: this.properties.trainingTitle,
        // trainingDocumentTitle: this.properties.trainingDocumentTitle
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
                }),
                // PropertyPaneTextField("trainingTitle", {
                //   label: "Training Title",
                //   value: this.properties.trainingTitle,
                //   placeholder: "Please enter training title."
                // }),
              ]
            }
          ]
        }
      ]
    };
  }
}
