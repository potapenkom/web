import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration
} from '@microsoft/sp-property-pane';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { sp } from '@pnp/sp';
import * as strings from 'SecafiIndicateursEtccWebPartStrings';
import SecafiIndicateursEtcc from './components/SecafiIndicateursEtcc';
import { ISecafiIndicateursEtccProps, ISecafiIndicateursEtccWebPartProps } from './components/ISecafiIndicateursEtccProps';



export default class SecafiIndicateursEtccWebPart extends BaseClientSideWebPart<ISecafiIndicateursEtccWebPartProps> {
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
    });
  }
  public render(): void {
    const element: React.ReactElement<ISecafiIndicateursEtccProps> = React.createElement(
      SecafiIndicateursEtcc,
      {
        description: this.properties.description,
        collectionData: this.properties.collectionData,
        context: this.context,
        fieldId: this.properties.fieldId,
        fieldTitle: this.properties.fieldTitle,
        listId: this.properties.listId,
        listTitle:this.properties.listTitle
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
                PropertyFieldCollectionData("collectionData", {
                  key: "collectionData",
                  label: "Collection data",
                  panelHeader: "Collection data panel header",
                  manageBtnLabel: "Manage collection data",
                  value: this.properties.collectionData,
                  fields: [
                    {
                      id: strings.listId,
                      title: strings.listTitle,
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: strings.fieldId,
                      title: strings.fieldTitle,
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                  ],
                  disabled: false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
