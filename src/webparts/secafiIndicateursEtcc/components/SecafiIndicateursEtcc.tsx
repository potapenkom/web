import * as React from 'react';
import styles from './SecafiIndicateursEtcc.module.scss';
import { ISecafiIndicateursEtccProps } from './ISecafiIndicateursEtccProps';

import { ISecafiIndicateursEtccState, ISearchRes } from './ISecafiIndicateursEtccState';
import { ChartControl, ChartType } from '@pnp/spfx-controls-react/lib/ChartControl';
import * as moment from 'moment'
import {
  DatePicker,
  defaultDatePickerStrings,
  Stack
} from '@fluentui/react';
import { sp, SearchQueryBuilder } from "@pnp/sp/presets/all";
import '@pnp/sp/search';
import { DefaultButton } from '@fluentui/react/lib/Button';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

const fileType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const fileExtension = '.xlsx';
let Heading1 = [["List name", "Field Name", "Field Value", "Année", "Produit", "Num mission", "Equipe", "Client", "Sortie de rapport"], []];
const saveExcel = (ListData, list1?, list2?) => {
  if (ListData.length > 0) {
    const ws = XLSX.utils.book_new();
    XLSX.utils.sheet_add_aoa(ws, Heading1);
    XLSX.utils.sheet_add_json(ws, ListData, { origin: 'A2', skipHeader: true });
    const wb = { Sheets: { 'Bilan_de_mission': ws }, SheetNames: ['Bilan_de_mission'] };
    if (list1.length > 0) {
      const ws = XLSX.utils.json_to_sheet(list1);
      XLSX.utils.sheet_add_aoa(ws, Heading1);
      XLSX.utils.book_append_sheet(wb, ws, "Suivi_de_relecture");
    }
    if (list2.length > 0) {
      const ws = XLSX.utils.json_to_sheet(list2);
      XLSX.utils.book_append_sheet(wb, ws, "Missions");
    }
    const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const data = new Blob([excelBuffer], { type: fileType });
    saveAs(data, 'Data' + fileExtension);
  }
}

export default class SecafiIndicateursEtcc extends React.Component<ISecafiIndicateursEtccProps, ISecafiIndicateursEtccState> {
  constructor(props) {
    super(props);
    this.state = {
      searchResults: [],
      searchPartRes: [],
      sortedResult: [],
      totalSearch: [],
      totalSearchBilan: [],
      totalSearchSuivi: [],
      totalSearchMissions: [],
      startDate: new Date('01-01-2021'),
      endDate: new Date('12-31-2021'),
    };
    this.GetData = this.GetData.bind(this);
  }

  public componentDidMount(): void {
    this.GetData();
  }

  public GetData(): void {
    this.setState({
      sortedResult: [],
      totalSearch: []
    });
    let start = moment(this.state.startDate).format('YYYY-MM-DD');
    let end = moment(this.state.endDate).format('YYYY-MM-DD');
    this.props.collectionData && this.props.collectionData.map(async (val) => {
      if (val.listId === "0x0100E297556C5DCE1F428F2CCB8A9A2609F6*") {
        let searchResults: ISearchRes[] = await this.searchBilan(val.listId, start, end, 0, 500, val.fieldId,);
        this.setState(prevState => ({ totalSearchBilan: prevState.totalSearchBilan.concat(searchResults) }))
        this.getPercent(searchResults);
      }
      if (val.listId === "0x01002229A785DC4FB442A6ABC3C478C38232*") {
        let searchResults: ISearchRes[] = await this.searchBilan('0x0100E297556C5DCE1F428F2CCB8A9A2609F6*', start, end, 0, 500);
        let serchSuinvi: ISearchRes[] = await this.searchSuiviRelecture(val.listId, val.fieldId, this.getMinMaxDate(searchResults, 'DDerniereReunion').minDate, this.getMinMaxDate(searchResults, 'DDerniereReunion').maxDate, 0, 500);
        this.setState(prevState => ({ totalSearchSuivi: prevState.totalSearchSuivi.concat(serchSuinvi) }))
        this.getPercent(this.hasBlanMission(searchResults, serchSuinvi));
      }
      if (val.listId === "0x010030F4365A045058449B6D5A1086834EB3007DA7964A5C6CE1479A322590C25A1CA5") {
        let maxDate = moment(start).subtract(1, 'M').format('YYYY-MM-DD');
        let minDate = moment(end).add(1, 'M').format('YYYY-MM-DD');
        let searchResults: ISearchRes[] = await this.searchBilan('0x0100E297556C5DCE1F428F2CCB8A9A2609F6*', start, end, 0, 500);
        let searchMission: ISearchRes[] = await this.searchMissions(val.listId, val.fieldId, maxDate, minDate, 0, 500);
        this.setState(prevState => ({ totalSearchMissions: prevState.totalSearchMissions.concat(this.hasBlanMissionSite(searchResults, searchMission)) }));
        this.getPercentMission(searchResults, searchMission);
      }
    })
  }

  //search Bilan_de_mission by ContentTypeID
  public searchBilan(cType: string, dStart: string, dEnd: string, row: number, pageSize: number, fieldName?: string): Promise<any> {
    let _results: ISearchRes[] = [];
    return new Promise((resolve, reject) => {
      const q =
        SearchQueryBuilder(`ContentTypeID:"${cType}"`)
          .selectProperties('DDerniereReunion', 'SPWebUrl', `${fieldName}`)
          .refinementFilters(`DDerniereReunion:range(${dStart},${dEnd})`)
          .startRow(row)
          .rowLimit(pageSize)
      sp.search(q).then((data) => {
        this.setState(prevState => ({ searchPartRes: prevState.searchPartRes.concat(data.PrimarySearchResults) }));
        let totalRows = data.TotalRows
        let nexstartRow = row + pageSize
        if (totalRows < nexstartRow) {
          this.state.searchPartRes.forEach(result => {
            _results.push({
              fieldValue: result[`${fieldName}`],
              DDerniereReunion: result['DDerniereReunion'],
              SPWebUrl: result['SPWebUrl'],
              listName: `${cType}`,
              fieldName: `${fieldName}`
            });
          });
          this.setState({
            searchPartRes: []
          })
          resolve(_results);
        } else {
          this.searchBilan(cType, dStart, dEnd, nexstartRow, pageSize, fieldName);
        }
      })
        .catch((ex) => {
          console.error(ex);
          reject(ex);
        });
    })
  }
  //search Suivi_de_relecture_par_relecteur by ContentTypeID
  public searchSuiviRelecture(cType: string, fieldName: string, dStart: string, dEnd: string, row: number, pageSize: number): Promise<any> {
    let _results: ISearchRes[] = [];
    return new Promise((resolve, reject) => {
      const q =
        SearchQueryBuilder(`ContentTypeID:"${cType}"`)
          .selectProperties(`${fieldName}`, 'Created', 'SPWebUrl')
          .refinementFilters(`Created:range(${dStart}, ${dEnd})`)
          .startRow(row)
          .rowLimit(pageSize)
      sp.search(q).then((data) => {
        let totalRows = data.TotalRows
        this.setState(prevState => ({ searchPartRes: prevState.searchPartRes.concat(data.PrimarySearchResults) }));
        let nexstartRow = row + pageSize
        if (totalRows < nexstartRow) {
          this.state.searchPartRes.forEach(result => {
            _results.push({
              fieldValue: result[`${fieldName}`],
              DCreation: result['Created'],
              SPWebUrl: result['SPWebUrl'],
              listName: `${cType}`,
              fieldName: `${fieldName}`
            });
          });
          this.setState({
            searchPartRes: []
          })
          resolve(_results);
        } else {
          this.searchSuiviRelecture(cType, fieldName, dStart, dEnd, nexstartRow, pageSize)
        }
      })
        .catch((ex) => {
          console.error(ex);
          reject(ex);
        });
    })
  }

  // search Missions by ContentTypeID
  public searchMissions(cType: string, fieldName: string, dStart: string, dEnd: string, row: number, pageSize: number): Promise<any> {
    let _results: ISearchRes[] = [];
    return new Promise((resolve, reject) => {
      const q =
        SearchQueryBuilder(`ContentTypeID:"${cType}"`)
          .selectProperties('Année', 'Produit', 'NumMission0', 'Equipe', 'Client', 'Sortie', 'SPWebUrl')
          .refinementFilters(`Sortie:range(${dStart}, ${dEnd})`)
          .startRow(row)
          .rowLimit(pageSize)
      sp.search(q).then((data) => {
        this.setState(prevState => ({ searchPartRes: prevState.searchPartRes.concat(data.PrimarySearchResults) }));
        let totalRows = data.TotalRows
        let nexstartRow = row + pageSize
        if (totalRows < nexstartRow) {
          this.state.searchPartRes.forEach(result => {
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
          });
          console.log('_results',_results)
          resolve(_results);
          this.setState({
            searchPartRes: []
          })
        } else {
          this.searchMissions(cType, fieldName, dStart, dEnd, nexstartRow, pageSize)
        }
      })
        .catch((ex) => {
          console.error(ex);
          reject(ex);
        });
    })
  }

  //data for exel export
  Listdata = () => {
    let resultBilan: ISearchRes[] = [];
    let resultSuivi: ISearchRes[] = [];
    let resultMision: ISearchRes[] = [];
    this.state.totalSearchBilan.forEach(element => {

      resultBilan.push({
        listName: this.decoderCType(element['listName']),
        fieldName: this.decoderField(element['fieldName']),
        fieldValue: element['fieldValue'],
        SPWebUrl: element['SPWebUrl']
      });
    })
    this.state.totalSearchSuivi.forEach(element => {
      resultSuivi.push({
        listName: this.decoderCType(element['listName']),
        fieldName: this.decoderField(element['fieldName']),
        fieldValue: element['fieldValue'],
        SPWebUrl: element['SPWebUrl']
      });
    })
    this.state.totalSearchMissions.forEach(element => {
      resultMision.push({
        listName: this.decoderCType(element['listName']),
        fieldName: this.decoderField(element['fieldName']),
        Annee: element['Annee'],
        Produit: element['Produit'],
        NumMission: element['NumMission'],
        Equipe: element['Equipe'],
        Client: element['Client'],
        Sortie: element['Sortie']
      });
    });

    //mapping Missions with Bilan_de_mission and Suivi_de_relecture_par_relecteur for export exel
    function getRelative(resultMision: ISearchRes[], resultItems: ISearchRes[]) {
      return resultItems.map(function (o1) {
        return resultMision.some(function (o2) {
          if (o2.NumMission) {
            var re = /-/gi;
            var NumMission = o2.NumMission.replace(re, "");
            let url = o1.SPWebUrl.split('/');
            let num = url.pop() || url.pop();
            if (NumMission === num) {
              resultItems.push({
                listName: o1['listName'],
                fieldName: o1['fieldName'],
                fieldValue: o1['fieldValue'],
                Annee: o2['Annee'],
                Produit: o2['Produit'],
                NumMission: o2['NumMission'],
                Equipe: o2['Equipe'],
                Client: o2['Client'],
                Sortie: o2['Sortie']
              })
            }
            delete o1.SPWebUrl;
          }
        });
      });
    }
    getRelative(resultMision, resultBilan);
    getRelative(resultMision, resultSuivi);
    saveExcel(resultBilan, resultSuivi, resultMision);
  }

  //mapping Bilan_de_mission with Suivi_de_relecture_par_relecteur
  hasBlanMission = (searchResults: ISearchRes[], serchSuinvi: ISearchRes[]): ISearchRes[] => {
    return serchSuinvi.filter(function (o1) {
      return searchResults.some(function (o2) {
        let urlSuivni = o1.SPWebUrl.split('/');
        let urlBilan = o2.SPWebUrl.split('/');
        let numSuivni = urlSuivni.pop() || urlSuivni.pop();
        let numBilan = urlBilan.pop() || urlBilan.pop();
        return numSuivni === numBilan; // return the ones with equal id
      });
    });
  }

  //Get dates for search Suivi_de_relecture_par_relecteur
  getMinMaxDate = (searchResults: ISearchRes[], dField: string) => {
    let rangeDate = {
      maxDate: '',
      minDate: ''
    }
    var minIdx = 0, maxIdx = 0;
    for (var i = 0; i < searchResults.length; i++) {
      if (searchResults[i][dField] > searchResults[maxIdx][dField]) maxIdx = i;
      if (searchResults[i][dField] < searchResults[minIdx][dField]) minIdx = i;
    }
    rangeDate.maxDate = moment(searchResults[maxIdx][dField]).add(1, 'M').format('YYYY-MM-DD');
    rangeDate.minDate = moment(searchResults[minIdx][dField]).subtract(1, 'M').format('YYYY-MM-DD');
    return rangeDate
  }

  //get Percent for Bilan_de_mission and Suivi_de_relecture_par_relecteur
  getPercent = (searchResults: ISearchRes[]) => {
    let countYes = 0;
    let countNo = 0;
    let listName;
    let fieldName;
    searchResults.forEach(element => {
      listName = element.listName;
      fieldName = element.fieldName
      if (element.fieldValue === "true" || element.fieldValue === "True\n\n1" || element.fieldValue === "OK" || element.fieldValue === "Oui") {
        countYes += 1;
      } else if (element.fieldValue === "false" || element.fieldValue === "False\n\n0" || element.fieldValue === "AR" || element.fieldValue === "Non") {
        countNo += 1;
      }
    });
    if ((countNo + countYes) > 0) {
      let arrayItem = {
        listId: listName,
        fieldId: fieldName,
        countYes: Math.round((100 * countYes) / (countYes + countNo)),
        countNo: Math.round((100 * countNo) / (countYes + countNo)),
      }
      let { sortedResult } = this.state;
      sortedResult.push(arrayItem);
      this.setState({
        sortedResult: sortedResult
      })
    }
  }

  //Chech has Miission Bilan_de_mission
  hasBlanMissionSite = (searchResults: ISearchRes[], searchMission: ISearchRes[]): ISearchRes[] => {
    return searchMission.filter(function (o1) {
      return searchResults.some(function (o2) {
        if (o1.NumMission) {
          var re = /-/gi;
          var NumMission = o1.NumMission.replace(re, "");
          let urlBilan = o2.SPWebUrl.split('/');
          let numBilan = urlBilan.pop() || urlBilan.pop();
          return NumMission === numBilan; // return the ones with equal id  
        }
      });
    });
  }

  //get Percent for Missions with Bilan_de_mission
  getPercentMission = (searchResults: ISearchRes[], searchMission: ISearchRes[]) => {
    let listName;
    let fieldName;
    let result = searchMission.filter(function (o1) {
      listName = o1.listName;
      fieldName = o1.fieldName;
      return searchResults.some(function (o2) {
        if (o1.NumMission) {
          var re = /-/gi;
          var NumMission = o1.NumMission.replace(re, "");
          let urlBilan = o2.SPWebUrl.split('/');
          let numBilan = urlBilan.pop() || urlBilan.pop();
          return NumMission === numBilan; // return the ones with equal id  
        }
      });
    });
    let arrayItem = {
      listId: listName,
      fieldId: fieldName,

      countYes: Math.round((100 * result.length) / (result.length + (searchMission.length - result.length))),
      countNo: Math.round((100 * (searchMission.length - result.length)) / (result.length + (searchMission.length - result.length))),
    }
    this.setState(prevState => ({ sortedResult: prevState.sortedResult.concat(arrayItem) }));

  }

  public render(): React.ReactElement<ISecafiIndicateursEtccProps> {
    return (
      <div className={styles.container}>
        <Stack horizontal>
          <DatePicker
            style={{ marginRight: '15px' }}
            label="Start Date"
            key={"dStart"}
            value={this.state.startDate}
            placeholder="Select start date..."
            isMonthPickerVisible={false}
            onSelectDate={date => this.setState({ startDate: date })}
            strings={defaultDatePickerStrings}
          />
          <DatePicker
            label="End Date"
            key={"dEnd"}
            value={this.state.endDate}
            placeholder="Select end date..."
            onSelectDate={date => this.setState({ endDate: date })}
            isMonthPickerVisible={false}
            strings={defaultDatePickerStrings}
          />
        </Stack>
        <DefaultButton id="Exel"
          onClick={this.Listdata}
          text="Export Exel"
          allowDisabledFocus
          style={{ marginRight: '27px' }} />
        <DefaultButton id="Refresh" onClick={this.GetData} text="Refresh data" allowDisabledFocus />
        <div className={styles.row}>
          {this.state.sortedResult.map((val) => {
            return (
              <ChartControl
                type={ChartType.Bar}
                data={{
                  labels: [`yes- ${val.countYes}%`, `no - ${val.countNo}%`],
                  datasets: [{
                    label: `${this.decoderCType(val.listId)} - ${this.decoderField(val.fieldId)}`,
                    data: [val.countYes, val.countNo, 0]
                  }]
                }} />
            );
          })}
        </div>
      </div>
    );
  }

//for display content type name
  decoderCType = (id: string) => {
    let cType: string
    switch (id) {
      case "0x010030F4365A045058449B6D5A1086834EB3007DA7964A5C6CE1479A322590C25A1CA5": {
        cType = "Mission";
        break;
      }
      case "0x01002229A785DC4FB442A6ABC3C478C38232*": {
        cType = "Suivi_de_relecture_par_relecteur";
        break;
      }
      case "0x0100E297556C5DCE1F428F2CCB8A9A2609F6*": {
        cType = "Bilan_de_mission";
        break;
      }
      default: {
        break;
      }
    }
    return cType
  }

//for display field name
  decoderField = (field: string) => {
    let fName: string
    switch (field) {
      case "Recommandations": {
        fName = "Recommandations (rapport)";
        break;
      }
      case "ReunionCadrageAvecDirection": {
        fName = "Réunion de cadrage avec la Direction";
        break;
      }
      case "ReunionPreparPleniereDirection": {
        fName = "Réunion préparatoire ou échanges avant la plénière avec la Direction";
        break;
      }
      case "RecueilSatisfactionCse": {
        fName = "Recueil formalisé de la satisfaction des élus du CSE";
        break;
      }
      case "PvCseRestitution": {
        fName = "PV du CSE de restitution récupéré et mis dans l’ETCC";
        break;
      }
      case field = "Sortie": {
        fName = "Sortie de rapport";
        break;
      }
      default: {
        break;
      }

    }
    return fName
  }
}