import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration
} from '@microsoft/sp-property-pane';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import pnp, { SearchResults, SearchQuery, SearchQueryBuilder } from "sp-pnp-js";
import "@pnp/sp/search";
import * as strings from 'SecafiIndicateursEtccWebPartStrings';
import SecafiIndicateursEtcc from './components/SecafiIndicateursEtcc';
import { ISecafiIndicateursEtccProps } from './components/ISecafiIndicateursEtccProps';

export interface ISecafiIndicateursEtccWebPartProps {
  description: string;
  collectionData: any[];
}

/*export function getSearchresults(cType: string, fieldName: string, sDate:Date, eDate:Date): Promise<any> {
  console.log('fieldName',fieldName)
  console.log('sDate',sDate)
  console.log('eDate',eDate)
  return new Promise((resolve, reject) => {

    let searchQuerySuivi : SearchQuery = {
      SelectProperties: [`${fieldName}`, 'Created', "SPWebUrl"],
    }

    let searchQueryBilan : SearchQuery = {
      SelectProperties: [`${fieldName}`, 'DDerniereReunion','SPWebUrl'],
     // RefinementFilters: [`DDerniereReunion:'range(2021-01-01, 2021-12-31)'`]
    }

    let searchQueryMissions: SearchQuery = {
      SelectProperties: ['Année', 'Produit', 'NumMission0', 'Equipe', 'Client', 'Sortie', 'SPWebUrl'],
    }
    let q;
    switch (cType) {
      case "Bilan_de_mission":
         q = SearchQueryBuilder.create(`ContentType:"${cType}"`, searchQueryBilan);
          break;
      case "Suivi_de_relecture_par_relecteur":
         q = SearchQueryBuilder.create(`ContentType:"${cType}"`, searchQuerySuivi);
          break;
      case "0x010030F4365A045058449B6D5A1086834EB3007DA7964A5C6CE1479A322590C25A1CA5":
        q = SearchQueryBuilder.create(`ContentTypeID:"${cType}"`, searchQueryMissions);
        break;
      default:
          console.log("No exists!");
          break;
  } 
  pnp.sp.search(q).then((r: SearchResults) => {
    resolve(r.PrimarySearchResults);
  })
    .catch((ex) => {
      console.error(ex);
      reject(ex);
    });
  });
}*/

export function getBilan(cType: string, fieldName:string, dStart:string, dEnd:string): Promise<any> {
  return new Promise((resolve, reject) => {
    pnp.sp.search(<SearchQuery>{
      Querytext: `ContentType:"${cType}"`,
      SelectProperties: [`${fieldName}`, 'DDerniereReunion', 'SPWebUrl'],
      RefinementFilters: [`DDerniereReunion:range(${dStart}, ${dEnd})`]
    }).then((r: SearchResults) => {
      resolve(r.PrimarySearchResults);
    })
      .catch((ex) => {
        console.error(ex);
        reject(ex);
      });
  });
}

export function getSuiviRelecture(cType: string, fieldName:string, dStart:string, dEnd:string): Promise<any> {
  return new Promise((resolve, reject) => {
    pnp.sp.search(<SearchQuery>{
      Querytext: `ContentType:"${cType}"`,
      SelectProperties: [`${fieldName}`, 'DDerniereReunion', 'SPWebUrl'],
      RefinementFilters: [`DDerniereReunion:range(${dStart}, ${dEnd})`]
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
