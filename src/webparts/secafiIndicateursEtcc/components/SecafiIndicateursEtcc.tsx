import * as React from 'react';
import styles from './SecafiIndicateursEtcc.module.scss';
import { ISecafiIndicateursEtccProps } from './ISecafiIndicateursEtccProps';
import { ISecafiIndicateursEtccState } from './ISecafiIndicateursEtccState';
import { ChartControl, ChartType } from '@pnp/spfx-controls-react/lib/ChartControl';
import { sp, IItemAddResult, DateTimeFieldFormatType } from "@pnp/sp/presets/all";
import * as moment from 'moment'
import {
  DatePicker,
  defaultDatePickerStrings,
  Stack
} from '@fluentui/react';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { getBilan } from '../SecafiIndicateursEtccWebPart';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

const fileType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const fileExtension = '.xlsx';
let Heading = [["Title", "Row"],];
const saveExcel = (ListData) => {
  if (ListData.length > 0) {
    const ws = XLSX.utils.book_new();
    // const ws = XLSX.utils.json_to_sheet(csvData,{header:["A","B","C","D","E","F","G"], skipHeader:false});  
    XLSX.utils.sheet_add_aoa(ws, Heading);
    XLSX.utils.sheet_add_json(ws, ListData, { origin: 'A2', skipHeader: true });
    const wb = { Sheets: { 'data': ws }, SheetNames: ['data'] };
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
      startDate: new Date('01-01-2021'),
      endDate: new Date('12-31-2021'),
    };
    sp.setup({
      spfxContext: this.props.context
    });
  }

  public componentDidMount(): void {
    let start = moment(this.state.startDate).format('YYYY-MM-DD');
    let end = moment(this.state.endDate).format('YYYY-MM-DD');

    this.props.collectionData && this.props.collectionData.map(async (val) => {
      if(val.listId ==="Bilan_de_mission"){
        let searchResults = await getBilan(val.listId, val.fieldId, start, end);
        console.log('searchResults from comp did mount', searchResults)
        let countYes = 0
        let countNo = 0
        searchResults.forEach(element => {
          if (element[val.fieldId] === "true" || element[val.fieldId] === "OK") {
            countYes += 1;
          } else if (element[val.fieldId] === "false" || element[val.fieldId] === "AR") {
            countNo += 1;
          }
        });
        if ((countNo + countYes) > 0) {
          let arrayItem = {
            listId: val.listId,
            fieldId: val.fieldId,
            countYes: (100 * countYes) / (countYes + countNo),
            countNo: (100 * countNo) / (countYes + countNo),
          }
          let { sortedResult } = this.state;
          sortedResult.push(arrayItem);
          this.setState({
            sortedResult: sortedResult
          })
        }  
      }
    })
  }

   Listdata = () => {
    saveExcel(this.state.searchResults);
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
            onSelectDate={this.handleStartDateSelection}
            strings={defaultDatePickerStrings}
          />
          <DatePicker
            label="End Date"
            key={"dEnd"}
            value={this.state.endDate}
            placeholder="Select end date..."
            onSelectDate={this.handleEndDateSelection}
            isMonthPickerVisible={false}
            strings={defaultDatePickerStrings}
          />
        </Stack>
        <DefaultButton id="Exel" onClick={this.Listdata} text="Export Exel" allowDisabledFocus />
        <DefaultButton id="Refresh" text="Refresh data" allowDisabledFocus />
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