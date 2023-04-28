import SPListViewService from "./listService/listService";
import { format } from 'date-fns'

const generateCSVdata = async (viewItemIds: string[], templateData: any[], templateType: string): Promise<any[]> => {

    const savedRows: string[] = [];

    const getSumma = (templateRow: string[], listItems: any[]) => {
        const csvName: string = templateRow[13];
        const currency: string = templateRow[8];
        let result: number = 0;
        if (savedRows.indexOf(csvName + currency) === -1) {
            listItems.forEach(item => {
                if (item.Affiliate.CSVName === csvName && item.Currency.Title === currency) {
                    if (templateRow[14].indexOf('betét') > -1) {
                        result = result + item.BankDeposit;
                    } else {
                        result = result + (item.CurrentAccounts + item.CashOnHandAndCheques);
                    }
                }
                templateRow[3] = format(new Date(item.DateOfBalance), 'yyyy.MM.dd');
                templateRow[4] = format(new Date(item.DateOfBalance), 'yyyy.MM.dd');
                templateRow[5] = format(new Date(item.DateOfBalance), 'yyyy.MM.dd');
            });
            savedRows.push(csvName + currency);
        }
        if (result === 0) {
            templateRow[7] = "1";
        } else {
            templateRow[7] = String(result);
        }
        return templateRow;
    }

    const getCurrencySumma = (templateRow: string[], listItems: any[]) => {
        const currency: string = templateRow[8];
        let result: number = 0;

        listItems.forEach(item => {
            if (item.Currency.Title === currency) {
                result = result + (item.CurrentAccounts + item.CashOnHandAndCheques + item.BankDeposit);
                
            }  
            templateRow[3] = format(new Date(item.DateOfBalance), 'yyyy.MM.dd');
            templateRow[4] = format(new Date(item.DateOfBalance), 'yyyy.MM.dd');
            templateRow[5] = format(new Date(item.DateOfBalance), 'yyyy.MM.dd');          
        });

        if (result === 0) {
            templateRow[7] = "1";
        } else {
            templateRow[7] = String(result);
        }
        return templateRow;
    }

    const listItems = await SPListViewService.getListItems(viewItemIds);
    //console.info(listItems);
    let result: any[] = [];
    switch (templateType) {
        case "offices":
            templateData.forEach((templateRow: string[], index: number) => {
                if (index === 0 && templateRow.length > 1) {
                    result.push(templateRow);
                } else if (templateRow.length > 1) {
                    result.push(getSumma(templateRow, listItems));
                }
            });
            break;
        case "subsidiaries":
            templateData.forEach((templateRow: string[], index: number) => {
                if (index === 0 && templateRow.length > 1) {
                    result.push(templateRow);
                } else if (templateRow.length > 1) {
                    result.push(getCurrencySumma(templateRow, listItems));
                }
            });
            break;
        default:
            break;
    }
    return result
}

export default generateCSVdata;