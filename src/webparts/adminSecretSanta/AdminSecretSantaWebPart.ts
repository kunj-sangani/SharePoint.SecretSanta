import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { sp } from "@pnp/sp";
import * as strings from 'AdminSecretSantaWebPartStrings';
import AdminSecretSanta from './components/AdminSecretSanta';
import { IAdminSecretSantaProps } from './components/IAdminSecretSantaProps';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';

export interface IAdminSecretSantaWebPartProps {
  description: string;
  adminUserEmail:string;
  lists: string;
}

export default class AdminSecretSantaWebPart extends BaseClientSideWebPart<IAdminSecretSantaWebPartProps> {
  
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
    });
  }

  public render(): void {
    const element: React.ReactElement<IAdminSecretSantaProps > = React.createElement(
      AdminSecretSanta,
      {
        description: this.properties.description,
        adminUserEmail:this.properties.adminUserEmail,
        context:this.context,
        lists: this.properties.lists
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
                  label: strings.DescriptionFieldLabel,
                  disabled:!(this.properties.adminUserEmail.toLowerCase()===this.context.pageContext.user.email.toLowerCase())
                }),
                PropertyPaneTextField('adminUserEmail', {
                  label: strings.adminUserEmailFieldLabel,
                  disabled:!(this.properties.adminUserEmail.toLowerCase()===this.context.pageContext.user.email.toLowerCase())
                }),
                PropertyFieldListPicker('lists', {
                  label: 'Select a list',
                  selectedList: this.properties.lists,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: !(this.properties.adminUserEmail.toLowerCase()===this.context.pageContext.user.email.toLowerCase()),
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
