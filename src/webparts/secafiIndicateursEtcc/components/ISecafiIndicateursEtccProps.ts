import {WebPartContext} from '@microsoft/sp-webpart-base';  

export interface ISecafiIndicateursEtccProps {
  description: string;
  collectionData: any[]; 
  context: WebPartContext;
  fieldId: string;  
  fieldTitle: string; 
  listId: string;  
  listTitle: string; 
}
export interface ISecafiIndicateursEtccWebPartProps {
  description: string;
  collectionData: any[];
  fieldId: string;  
  fieldTitle: string; 
  listId: string;  
  listTitle: string; 

}
