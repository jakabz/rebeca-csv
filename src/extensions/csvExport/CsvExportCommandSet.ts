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
import { CSVLink } from "react-csv";
import * as ReactDom from 'react-dom';
import { format } from 'date-fns';
import * as iconv from 'iconv-lite';

let currentViewId: string = '';

export interface IEnableListView {
  viewId: string;
  spGroupsTitle: string[];
  templateType: 'offices' | 'subsidiaries';
  templateCSV: string;
}

export interface ICsvExportCommandSetProperties {
  enablelistview: IEnableListView[];
}

export default class CsvExportCommandSet extends BaseListViewCommandSet<ICsvExportCommandSetProperties> {

  public async onInit(): Promise<void> {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    currentViewId = this.context.listView.view.id.toString();
    let showInView: boolean = false;
    for (const listView of this.properties.enablelistview) {
      const isMember: boolean = await SPListViewService.isGroupMember(listView.spGroupsTitle);
      if (isMember) {
        if (listView.viewId.toUpperCase() === currentViewId.toUpperCase()) {
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
    for (const listView of this.properties.enablelistview) {
      const isMember: boolean = await SPListViewService.isGroupMember(listView.spGroupsTitle);
      if (isMember) {
        if (listView.viewId.toUpperCase() === currentViewId.toUpperCase()) {
          showInView = true;
        }
      }
    }
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    compareOneCommand.visible = compareOneCommand && showInView;

    this.raiseOnChange();
  }

  public fetchCSV = async (file: string): Promise<string> => {
    const response = await fetch(this.context.pageContext.legacyPageContext.webAbsoluteUrl + '/SiteAssets/' + file, {
      credentials: 'include',
      headers: {
        contentType: "application/json; charset=utf-8",
      }
    });
    const data = await response.text();
    return data;
  }

  public processCSV = (str: string, delim = ','): string[][] => {
    const rows = str.split("\n");
    const arr = rows.map(function (row) {
      const values = row.split(delim);
      return values;
    });
    return arr;
  }

  private csvExportClicks = (): void => {
    const viewItems: readonly RowAccessor[] = this.context.listView.rows;
    const viewItemIds = viewItems.map(row => row.getValueByName('ID'));
    const settingViewItem = this.properties.enablelistview.find(p => p.viewId.toUpperCase() === currentViewId.toUpperCase());
    const fileName: string = `${settingViewItem.templateType}-${format(new Date(), 'yyyy.MM.dd.hh:mm:ss')}.csv`;

    this.fetchCSV(settingViewItem.templateCSV)
      .then(async result => {
        const templateData = this.processCSV(result, ';');
        const csvData = await generateCSVdata(viewItemIds, templateData, settingViewItem.templateType);

        const element = React.createElement(CSVLink, {
          data: csvData,
          filename: fileName,
          separator: ";",
          enclosingCharacter: ``,
          uFEFF: false
        }, 'download');

        const downloadContainer: HTMLElement = document.createElement('div');
        downloadContainer.setAttribute('id', 'downloadContainer');
        ReactDom.render(element, downloadContainer);

        // (downloadContainer.querySelector("a") as HTMLElement).click();
        await this.downloadCsvAsANSI(downloadContainer, fileName);

        // Unmount component to avoid memory leaks.
        ReactDom.unmountComponentAtNode(downloadContainer);

      })
      .catch(error => console.error('Error (fetchCSV): ', error));
  }

  private downloadCsvAsANSI = (container: HTMLElement, fileName: string): void => {
    const bloblUrl: string = (container.querySelector('a') as HTMLAnchorElement)?.href;
    fetch(bloblUrl)
      .then(async (result) => {
        let resultText: string = await result.text();
        resultText = iconv.encode(resultText, 'win1252');

        const filData = new Blob([resultText], { type: 'text/csv' });
        const filUrl = URL.createObjectURL(filData);

        const link: HTMLAnchorElement = document.createElement("a");
        link.href = filUrl;
        link.download = fileName;
        link.click();
      })
      .catch(error => console.error('Error (downloadCsvAsANSI): ', error));
  }
}
