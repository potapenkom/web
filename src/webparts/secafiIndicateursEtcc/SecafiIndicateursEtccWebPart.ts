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

export interface ISearchResult {
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

export interface ISearchBilan extends ISearchResult {
  DDerniereReunion: Date;
}

export interface ISearchSuivi extends ISearchResult {
  DCreation: Date
}

export interface ISearchMissions extends ISearchResult {
  Sortie: Date;
  Annee: string;
  Produit: string;
  NumMission: string;
  Equipe: string;
  Client: string;
}

export function getBilan(cType: string, dStart: string, dEnd: string, fieldName?: string): Promise<any> {
  const _results: ISearchBilan[] = [];
  return new Promise((resolve, reject) => {
    pnp.sp.search(<SearchQuery>{
      Querytext: `ContentTypeID:"0x0100E297556C5DCE1F428F2CCB8A9A2609F6*"`,
      RowLimit: 9999,
      SelectProperties: ['DDerniereReunion', 'SPWebUrl', `${fieldName}`],
      RefinementFilters: [`DDerniereReunion:range(${dStart},${dEnd})`]
    }).then((r: SearchResults) => {
      r.PrimarySearchResults.forEach(result => {
        _results.push({
          fieldValue: result[`${fieldName}`],
          DDerniereReunion: result['DDerniereReunion'],
          SPWebUrl: result['SPWebUrl'],
          listName: `${cType}`,
          fieldName: `${fieldName}`
        });
      })
      resolve(_results);
    })
      .catch((ex) => {
        console.error(ex);
        reject(ex);
      });
  });
}

export function getSuiviRelecture(cType: string, fieldName: string, dStart: string, dEnd: string): Promise<any> {
  const _results: ISearchSuivi[] = [];
  return new Promise((resolve, reject) => {
    pnp.sp.search(<SearchQuery>{
      Querytext: `ContentTypeID:"${cType}"`,
      RowLimit: 9999,
      SelectProperties: [`${fieldName}`, 'Created', 'SPWebUrl'],
      RefinementFilters: [`Created:range(${dStart}, ${dEnd})`]
    }).then((r: SearchResults) => {
      r.PrimarySearchResults.forEach(result => {
        _results.push({
          fieldValue: result[`${fieldName}`],
          DCreation: result['Created'],
          SPWebUrl: result['SPWebUrl'],
          listName: `${cType}`,
          fieldName: `${fieldName}`
        });
      })
      resolve(_results);
    })
      .catch((ex) => {
        console.error(ex);
        reject(ex);
      });
  });
}

export function getMissions(cType: string, fieldName: string, dStart: string, dEnd: string): Promise<any> {
  const _results: ISearchMissions[] = [];
  return new Promise((resolve, reject) => {
    pnp.sp.search(<SearchQuery>{
      Querytext: `ContentTypeID:"${cType}"`,
      RowLimit: 9999,
      SelectProperties: ['Année', 'Produit', 'NumMission0', 'Equipe', 'Client', 'Sortie', 'SPWebUrl'],
      RefinementFilters: [`Sortie:range(${dStart}, ${dEnd})`],
    }).then((r: SearchResults) => {
      let totalRows: number = r.TotalRows;
      let pageSize: number = 500
      r.PrimarySearchResults.forEach(result => {
        _results.push({
          fieldValue: result[`${fieldName}`],
          SPWebUrl: result['SPWebUrl'],
          listName: `${cType}`,
          fieldName: `${fieldName}`,
          Sortie: result['Sortie'],
          Annee: result['Année'],
          Produit: result['Produit'],
          NumMission: result['NumMission0'],
          Equipe: result['Equipe'],
          Client: result['Client'],
        });
      })
      resolve(_results);
    })
      .catch((ex) => {
        console.error(ex);
        reject(ex);
      });
  });
}

export function getFullMissions(cType: string, fieldName: string, dStart: string, dEnd: string): Promise<any> {
  const _results: ISearchMissions[] = [];
  let currentResults: any[] = null;
  return new Promise((resolve, reject) => {
    startSearch();
    function startSearch(page: number = 0) {
      pnp.sp.search(<SearchQuery>{
        Querytext: `ContentTypeID:"${cType}"`,
        StartRow: page,
        RowLimit: 500,
        SelectProperties: ['Année', 'Produit', 'NumMission0', 'Equipe', 'Client', 'Sortie', 'SPWebUrl'],
        // RefinementFilters: [`Sortie:range(${dStart}, ${dEnd})`],
      }).then((r: SearchResults) => {
        let totalRows: number = r.TotalRows;
        console.log('totalRows',totalRows);
        let pageSize: number = 500
        currentResults = r.PrimarySearchResults;
        console.log('currentResults',currentResults)
        if (totalRows > pageSize) {
          let totalPages = parseInt((totalRows / pageSize).toString());
          console.log('totalPages',totalPages);
          for (let page = 1; page <= totalPages; page++) {
            let startRow = page * pageSize;
            console.log('startRow',startRow)
           // startSearch(startRow);
          }
        }
        currentResults.forEach(result => {
          _results.push({
            fieldValue: result[`${fieldName}`],
            SPWebUrl: result['SPWebUrl'],
            listName: `${cType}`,
            fieldName: `${fieldName}`,
            Sortie: result['Sortie'],
            Annee: result['Année'],
            Produit: result['Produit'],
            NumMission: result['NumMission0'],
            Equipe: result['Equipe'],
            Client: result['Client'],
          });
        })
        resolve(_results);
      })
        .catch((ex) => {
          console.error(ex);
          reject(ex);
        });
    }
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
