import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SpFxListSearchWebPartWebPartStrings';
import SpFxListSearchWebPart from './components/SpFxListSearchWebPart';
import { ISpFxListSearchWebPartProps } from './components/ISpFxListSearchWebPartProps';
import { getSP } from './pnpjsConfig';
import { SPService } from './Services/Service';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy, PropertyFieldMultiSelect } from '@pnp/spfx-property-controls';

export interface ISelectedListProps {
  id: string;
  title: string;
  url: string;
}

export interface ISpFxListSearchWebPartWebPartProps {
  listFields: IPropertyPaneDropdownOption[];
  spListType: string;
  fields: any[];
  columnsToShow: any[];
  lists: ISelectedListProps;
  description: string;
}

export default class SpFxListSearchWebPartWebPart extends BaseClientSideWebPart<ISpFxListSearchWebPartWebPartProps> {

  private _services : SPService;
  private _listFields: IPropertyPaneDropdownOption[] = [];

  public render(): void {
    const element: React.ReactElement<ISpFxListSearchWebPartProps> = React.createElement(
      SpFxListSearchWebPart,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
      await super.onInit();

      getSP(this.context);
      this._services = new SPService();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private getListFields = async () => {
    if(this.properties.lists) {
      let allFields: any = await this._services.getFields(this.properties.lists.id);

      this._listFields = [];
      this._listFields.push(...allFields.map(field => ( { key: field.InternalName, text: field.Title, fieldType: field['odata.type'] } )));
      this.properties.listFields = this._listFields;
      this.properties.description = this.properties.lists.title;

      //Check if Selected List is SP List or SP Document Library
      const listTypeArr = this.properties.listFields.filter(field => field.key === "FileLeafRef");
      if(listTypeArr.length > 0) {
        this.properties.spListType = "SPDocumentLibrary";
      } else {
        this.properties.spListType = "SPList";
      }
    }
  }

  private listConfigurationChanged = async (propertyPath: string, oldValue: any, newValue: any) => {
    await this.getListFields();

    if(propertyPath === 'lists' && newValue) {
      this.properties.fields = [];
      this.properties.columnsToShow = [];

      this.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    } else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    }
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
                PropertyFieldListPicker('lists', {
                  label: "Select a List",
                  selectedList: this.properties.lists,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  properties: this.properties,
                  context: this.context as any,
                  key: "listPickerFieldId",
                  includeListTitleAndUrl: true,
                  onPropertyChange: this.listConfigurationChanged,
                }),
                PropertyFieldMultiSelect('fields', {
                  key: "filterFields",
                  label: "Select Fields to Filter",
                  options: this.properties.listFields,
                  selectedKeys: this.properties.fields,
                }),
                PropertyFieldMultiSelect('columnsToShow', {
                  key: 'displayFields',
                  label: "Select Fields to Display",
                  options: this.properties.listFields,
                  selectedKeys: this.properties.columnsToShow,
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
