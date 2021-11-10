import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration
} from '@microsoft/sp-property-pane';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import pnp, { SearchResults, SearchQuery, SearchQueryBuilder } from "sp-pnp-js";
import * as strings from 'SecafiIndicateursEtccWebPartStrings';
import SecafiIndicateursEtcc from './components/SecafiIndicateursEtcc';
import { ISecafiIndicateursEtccProps } from './components/ISecafiIndicateursEtccProps';

export interface ISecafiIndicateursEtccWebPartProps {
  description: string;
  collectionData: any[];
}

export function getSearchresults(cType: string, fieldName: string): Promise<any> {
  return new Promise((resolve, reject) => {
    pnp.sp.search(<SearchQuery>{
      Querytext: `ContentType: ${cType}`,
      SelectProperties: [`${fieldName}`, 'DDerniereReunion'],
    }).then((r: SearchResults) => {
      resolve(r.PrimarySearchResults);
    })
      .catch((ex) => {
        console.error(ex);
        reject(ex);
      });
  });
}


export default class SecafiIndicateursEtccWebPart extends BaseClientSideWebPart<ISecafiIndicateursEtccWebPartProps> {

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
                      title: "List",
                      type: CustomCollectionFieldType.dropdown,
                      options: [
                        {
                          key: "Suivi de relecture",
                          text: "Suivi de relecture"
                        },
                        {
                          key: "Bilan de mission",
                          text: "Bilan de mission"
                        },
                        {
                          key: "Missions",
                          text: "Missions"
                        },
                      ],
                      required: true
                    },

                    {
                      id: "fieldId",
                      title: "Field",
                      type: CustomCollectionFieldType.dropdown,
                      options: [
                        {
                          key: "Recommandations",
                          text: "Recommandations (rapport)"
                        },
                        {
                          key: "ReunionCadrageAvecDirection",
                          text: "Réunion de cadrage avec la Direction"
                        },
                        {
                          key: "ReunionPreparPleniereDirection",
                          text: "Réunion préparatoire ou échanges avant la plénière avec la Direction"
                        },
                        {
                          key: "RecueilSatisfactionCse",
                          text: "Recueil formalisé de la satisfaction des élus du CSE"
                        },
                        {
                          key: "PvCseRestitution",
                          text: "PV du CSE de restitution récupéré et mis dans l’ETCC"
                        },
                        {
                          key: "SortieRapport",
                          text: "Sortie de rapport"
                        },
                      ],
                      required: true
                    }
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
