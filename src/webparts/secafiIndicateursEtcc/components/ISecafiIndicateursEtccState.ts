import {  ISearchResult } from "@pnp/sp/presets/all";
export interface ISecafiIndicateursEtccState {
  searchResults: ISearchRes[];
  searchPartRes : ISearchResult[];
  sortedResult: any[];
  totalSearch: any[];
  totalSearchBilan: any[];
  totalSearchSuivi: any[];
  totalSearchMissions: any[];
  startDate: Date;
  endDate: Date;
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



