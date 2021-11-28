import * as React from 'react';
import styles from './SecafiIndicateursEtcc.module.scss';
import { ISecafiIndicateursEtccProps } from './ISecafiIndicateursEtccProps';
import { ISecafiIndicateursEtccState } from './ISecafiIndicateursEtccState';
import { ChartControl, ChartType } from '@pnp/spfx-controls-react/lib/ChartControl';
import { sp } from "@pnp/sp/presets/all";
import * as moment from 'moment'
import {
  DatePicker,
  defaultDatePickerStrings,
  Stack
} from '@fluentui/react';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { getBilan, ISearchBilan, getSuiviRelecture, ISearchResult, getMissions, ISearchMissions } from '../SecafiIndicateursEtccWebPart';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import { times } from 'lodash';

const fileType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const fileExtension = '.xlsx';
let Heading1 = [["List name", "Field Name", "Field Value", "Année", "Produit", "Num mission", "Equipe", "Client", "Sortie de rapport"],[]];
let Heading2 = [["List name","Année", "Produit", "Num mission", "Equipe", "Client", "Sortie de rapport"],[]];

const saveExcel = (ListData) => {
  if (ListData.length > 0) {
    const ws = XLSX.utils.book_new();
    XLSX.utils.sheet_add_aoa(ws, Heading1);
    XLSX.utils.sheet_add_json(ws, ListData, { origin: 'A2', skipHeader: true });
    const wb = { Sheets: { 'Bilan_de_mission': ws }, SheetNames: ['Bilan_de_mission'] };
    XLSX.utils.book_append_sheet(wb, ws, "Suivi_de_relecture");
    XLSX.utils.sheet_add_aoa(ws, Heading2);
    XLSX.utils.book_append_sheet(wb, ws, "Missions");
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
      sortedResult: [],
      totalSearch:[],
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
      totalSearch:[]
    });
    let start = moment(this.state.startDate).format('YYYY-MM-DD');
    let end = moment(this.state.endDate).format('YYYY-MM-DD');
    this.props.collectionData && this.props.collectionData.map(async (val) => {
      if (val.listId === "Bilan_de_mission") {
        let searchResults: ISearchResult[] = await getBilan(val.listId, start, end, val.fieldId,);
        this.setState(prevState => ({ totalSearch: prevState.totalSearch.concat(searchResults) }))
        console.log(this.state.totalSearch)
        console.log('searchResults Bilan_de_mission', searchResults)
        this.getPercent(searchResults);
      }
      if (val.listId === "Suivi_de_relecture_par_relecteur") {
        let searchResults: ISearchResult[] = await getBilan('Bilan_de_mission', start, end);
        console.log('getBilan', searchResults)
        let serchSuinvi: ISearchResult[] = await getSuiviRelecture(val.listId, val.fieldId, this.getMinMaxDate(searchResults, 'DDerniereReunion').minDate, this.getMinMaxDate(searchResults, 'DDerniereReunion').maxDate);
        console.log('searchResults Suivi_de_relecture_par_relecteur', serchSuinvi);
        this.setState(prevState => ({ totalSearch: prevState.totalSearch.concat(serchSuinvi) }))
        console.log(this.state.totalSearch)

        this.getPercent(this.hasBlanMission(searchResults, serchSuinvi));
      }
      if (val.listId === "0x010030F4365A045058449B6D5A1086834EB3007DA7964A5C6CE1479A322590C25A1CA5") {
        let maxDate = moment(start).subtract(1, 'M').format('YYYY-MM-DD');
        let minDate = moment(end).add(1, 'M').format('YYYY-MM-DD');

        let searchResults: ISearchResult[] = await getBilan('Bilan_de_mission', start, end);
        let searchMission: ISearchMissions[] = await getMissions(val.listId, val.fieldId, maxDate, minDate);
        console.log('searchMission ', searchMission)
        this.getPercentMission(searchResults, searchMission);
      }
    })

  }

  Listdata = () => {
    saveExcel(this.state.sortedResult);
  }

  hasBlanMission = (searchResults: ISearchResult[], serchSuinvi: ISearchResult[]): ISearchResult[] => {
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

  getMinMaxDate = (searchResults: ISearchResult[], dField: string) => {
    let rangeDate = {
      maxDate: '',
      minDate: ''
    }
    var minIdx = 0, maxIdx = 0;
    for (var i = 0; i < searchResults.length; i++) {
      if (searchResults[i][dField] > searchResults[maxIdx][dField]) maxIdx = i;
      if (searchResults[i][dField] < searchResults[minIdx][dField]) minIdx = i;
    }
    console.log( 'max',searchResults[maxIdx][dField])
    console.log( 'min',searchResults[minIdx][dField])

    rangeDate.maxDate = moment(searchResults[maxIdx][dField]).add(1, 'M').format('YYYY-MM-DD');
    rangeDate.minDate = moment(searchResults[minIdx][dField]).subtract(1, 'M').format('YYYY-MM-DD');
    return rangeDate
  }


  getPercent = (searchResults: ISearchResult[]) => {
    let countYes = 0;
    let countNo = 0;
    let listName;
    let fieldName;
    searchResults.forEach(element => {
      listName = element.listName;
      fieldName = element.fieldName
      if (element.fieldValue === "true" || element.fieldValue === "OK") {
        countYes += 1;
      } else if (element.fieldValue === "false" || element.fieldValue === "AR") {
        countNo += 1;
      }
    });
    if ((countNo + countYes) > 0) {
      let arrayItem = {
        listId: listName,
        fieldId: fieldName,
        countYes: Math.round((100 * countYes) / (countYes + countNo)) ,
        countNo: Math.round((100 * countNo) / (countYes + countNo)),
      }
      let { sortedResult } = this.state;
      sortedResult.push(arrayItem);
      this.setState({
        sortedResult: sortedResult
      })
    }
  }

  getPercentMission = (searchResults: ISearchResult[], searchMission: ISearchMissions[]) => {
    let result = searchMission.filter(function (o1) {
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
      listId: "Mission",
      fieldId: "Sortie de rapport",

      countYes: Math.round((100 * result.length) / (result.length + (searchMission.length - result.length))),
      countNo:  Math.round((100 *(searchMission.length - result.length)) / (result.length + (searchMission.length - result.length))),
    }

    let { sortedResult } = this.state;
    sortedResult.push(arrayItem);
    this.setState({
      sortedResult: sortedResult
    })
  }

  public render(): React.ReactElement<ISecafiIndicateursEtccProps> {
    return (
      <div className={styles.container}>
        <Stack horizontal>
          <DatePicker
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
        <DefaultButton id="Exel" onClick={this.Listdata} text="Export Exel" allowDisabledFocus />
        <DefaultButton id="Refresh" onClick={this.GetData} text="Refresh data" allowDisabledFocus />
        <div className={styles.row}>
          {this.state.sortedResult.map((val) => {
            console.log('map sortedResult ', this.state.sortedResult)
            return (
              <ChartControl
                type={ChartType.Bar}
                data={{
                  labels: [`yes- ${val.countYes}%`, `no - ${val.countNo}%`],
                  datasets: [{
                    label: `${val.listId} - ${val.fieldId}`,
                    data: [val.countYes, val.countNo, 0]
                  }]
                }} />
            );
          })}
        </div>
      </div>
    );
  }

  handleStartDateSelection = (date: Date | null | undefined): void => {
    this.setState({ startDate: date });
  }
  handleEndDateSelection = (date: Date | null | undefined): void => {
    this.setState({ endDate: date });
  }

  decoderCType = (id: string) => {
    let cType: string
    switch (id) {
      case "0x010030F4365A045058449B6D5A1086834EB3007DA7964A5C6CE1479A322590C25A1CA5": {
        cType = "Mission";
        break;
      }
      case "0x01002229A785DC4FB442A6ABC3C478C38232": {
        cType = "Suivi_de_relecture_par_relecteur";
        break;
      }
      case "0x0100E297556C5DCE1F428F2CCB8A9A2609F6": {
        cType = "Bilan_de_mission";
        break;
      }
      default: {
        break;
      }
    }
    console.log("cType", cType)

    return cType
  }

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
    console.log("fName", fName)
    return fName
  }
}