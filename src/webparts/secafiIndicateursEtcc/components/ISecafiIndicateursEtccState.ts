import {  ISearchResult } from "@pnp/sp/presets/all";
import { ISearchRes} from '../SecafiIndicateursEtccWebPart'
export interface ISecafiIndicateursEtccState {
  searchResults: ISearchRes[];
  searchPartRes : ISearchResult[]
  sortedResult: any[];
  totalSearch: any[];
  totalSearchBilan: any[],
  totalSearchSuivi: any[],
  totalSearchMissions: any[],
  startDate: Date;
  endDate: Date;
}


