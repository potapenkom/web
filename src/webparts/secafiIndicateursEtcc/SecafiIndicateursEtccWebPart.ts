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
import { ISecafiIndicateursEtccProps } from './components/ISecafiIndicateursEtccProps';

export interface ISecafiIndicateursEtccWebPartProps {
  description: string;
  collectionData: any[];
}

export interface ISearchRes {
  listName: string;
  fieldName: string;
  fieldValue?: string;
  SPWebUrl?: string;
  DDerniereReunion?: Date;
  DCreation?: Date;
  Sortie?: Date;
  Annee?: string;
  Produit?: string;
  NumMission?: string;
  Equipe?: string;
  Client?: string;
}

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
                PropertyFieldCollectionData("collectionData", {
                  key: "collectionData",
                  label: "Collection data",
                  panelHeader: "Collection data panel header",
                  manageBtnLabel: "Manage collection data",
                  value: this.properties.collectionData,
                  fields: [
                    {
                      id: "listId",
                      title: "Content Type ID",
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: "fieldId",
                      title: "Column internal name",
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
