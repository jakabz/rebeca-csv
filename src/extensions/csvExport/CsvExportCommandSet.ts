import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetExecuteEventParameters,
  ListViewStateChangedEventArgs,
  RowAccessor
} from '@microsoft/sp-listview-extensibility';
import * as React from 'react';
import generateCSVdata from './generateCSVdata';
import SPListViewService from './listService/listService';
import { CSVLink, CSVDownload } from "react-csv";
import * as ReactDom from 'react-dom';
import { format } from 'date-fns';


let currentViewId: string = '';

export interface ICsvExportCommandSetProperties {
  enablelistview: any[];
  viewItems: any[];
}

export default class CsvExportCommandSet extends BaseListViewCommandSet<ICsvExportCommandSetProperties> {

  public async onInit(): Promise<void> {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    currentViewId = this.context.listView.view.id.toString();
    let showInView: boolean = false;
    for (var index in this.properties.enablelistview) {
      const isMember: boolean = await SPListViewService.isGroupMember(this.properties.enablelistview[index].spGroupsTitle);
      if (isMember) {
        if (this.properties.enablelistview[index].viewId.toUpperCase() === currentViewId.toUpperCase()) {
          showInView = true;
        }
      }
    }
    compareOneCommand.visible = compareOneCommand && showInView;

    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_1':
        this.csvExportClicks();
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private _onListViewStateChanged = async (args: ListViewStateChangedEventArgs): Promise<void> => {
    currentViewId = this.context.listView.view.id.toString();
    let showInView: boolean = false;
    for (var index in this.properties.enablelistview) {
      const isMember: boolean = await SPListViewService.isGroupMember(this.properties.enablelistview[index].spGroupsTitle);
      if (isMember) {
        if (this.properties.enablelistview[index].viewId.toUpperCase() === currentViewId.toUpperCase()) {
          showInView = true;
        }
      }
    }
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    compareOneCommand.visible = compareOneCommand && showInView;

    this.raiseOnChange();
  }

  public fetchCSV = async (file: string) => {
    const response = await fetch(this.context.pageContext.legacyPageContext.webAbsoluteUrl + '/SiteAssets/' + file, {
      credentials: 'include',
      headers: {
        contentType: "application/json; charset=utf-8",
      }
    });
    const data = await response.text();
    return data;
  }

  public processCSV = (str: string, delim = ',') => {
    const rows = str.split("\n");
    const arr = rows.map(function (row) {
      const values = row.split(delim);
      return values;
    });
    return arr;
  }

  private csvExportClicks = () => {
    const viewItems: readonly RowAccessor[] = this.context.listView.rows;
    const viewItemIds = viewItems.map(row => row.getValueByName('ID'));
    const settingViewItem = this.properties.enablelistview.find(p => p.viewId.toUpperCase() === currentViewId.toUpperCase());
    this.fetchCSV(settingViewItem.templateCSV).then(async result => {
      const templateData = this.processCSV(result, ';');
      const csvData = await generateCSVdata(viewItemIds, templateData, settingViewItem.templateType);

      const element = React.createElement(CSVLink, {
        data: csvData,
        filename: `${settingViewItem.templateType}-${format(new Date(), 'yyyy.MM.dd.hh:mm:ss')}.csv`,
        separator: ";",
        enclosingCharacter: ``,
        uFEFF: false
      }, 'download');
      let downloadContainer: HTMLElement = document.createElement('div');
      downloadContainer.setAttribute('id', 'downloadContainer');
      ReactDom.render(element, downloadContainer);

      (downloadContainer.querySelector("a") as HTMLElement).click();

      //console.info(csvData);
    });
  }
}
