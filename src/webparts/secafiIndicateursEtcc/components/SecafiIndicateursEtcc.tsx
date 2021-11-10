import * as React from 'react';
import styles from './SecafiIndicateursEtcc.module.scss';
import { ISecafiIndicateursEtccProps } from './ISecafiIndicateursEtccProps';
import { ISecafiIndicateursEtccState } from './ISecafiIndicateursEtccState';
import { ChartControl, ChartType } from '@pnp/spfx-controls-react/lib/ChartControl';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  DatePicker,
  defaultDatePickerStrings,
  Stack
} from '@fluentui/react';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { getSearchresults } from '../SecafiIndicateursEtccWebPart';

export default class SecafiIndicateursEtcc extends React.Component<ISecafiIndicateursEtccProps, ISecafiIndicateursEtccState> {
  constructor(props) {
    super(props);
    this.state = {
      searchResults: [],
      startDate: new Date('01-01-2021'),
      endDate: new Date('12-31-2021'),
      percentageTrue: 50,
      percentageFalse: 50
    };
  }

  public componentDidMount(): void {
    this.props.collectionData && this.props.collectionData.map(async (val) => {

      let searchResults = await getSearchresults(val.listId, val.fieldId);
      console.log("searchResults", searchResults)
      searchResults.map((result) => {
        this.setState({ searchResults: [...this.state.searchResults, result] })
      })
    })
  }


  public render(): React.ReactElement<ISecafiIndicateursEtccProps> {
    let filterState = this._filterDate(this.state.searchResults)
    let arr: number[] = [1, 0, 1, 0, 1, 0, 1, 0, 1, 1, 1];
    console.log("state searchResults ", this.state.searchResults)
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
            // DatePicker uses English strings by default. For localized apps, you must override this prop.
            strings={defaultDatePickerStrings}
          />
          <DatePicker
            label="End Date"
            key={"dEnd"}
            value={this.state.endDate}
            placeholder="Select end date..."
            onSelectDate={this.handleEndDateSelection}
            isMonthPickerVisible={false}
            // DatePicker uses English strings by default. For localized apps, you must override this prop.
            strings={defaultDatePickerStrings}
          />
        </Stack>

        <DefaultButton text="Export Exel" allowDisabledFocus />
        <h1>Props</h1>
        <div className={styles.row}>
          {this.props.collectionData && this.props.collectionData.map((val) => {
            console.log('collectionData ', this.props.collectionData)
            return (
              <ChartControl
                type={ChartType.Bar}
                data={{
                  labels: [`${val.fieldId}`, `${val.fieldId}`],
                  datasets: [{
                    label: `${val.listId}`,
                    data: [60, 40, 0]
                  }]
                }} />
            );
          })}
        </div>
      </div>
    );
  }
  private _filterDate(state: any[]) {
    return state.filter(item => {
      let d = item['DDerniereReunionOWSDATE']
      console.log("d", d)
      let date = new Date(d);
      console.log("date", date)
      return date >= this.state.startDate && date <= this.state.endDate;
    }
    )
  }

  handleStartDateSelection = (e) => {
    this.setState({ startDate: e });
  }

  handleEndDateSelection = (e) => {
    this.setState({ endDate: e });
  }

  percentageTrue = (state: any[]) => {
    let partialValue = state.filter(num => num === 1).length;
    let totalValue = state.length;
    var result = (100 * partialValue) / totalValue;
    console.log('percentageTrue ',result)
    this.setState({ percentageTrue: result });
    return result;
  }
  percentageFalse = (state: any[]) => {
    let partialValue = state.filter(num => num === 0).length;
    let totalValue = state.length;
    var result = (100 * partialValue) / totalValue;
    console.log('percentageFalse ',result)
    this.setState({ percentageFalse: result });
    return result;
  }

}

