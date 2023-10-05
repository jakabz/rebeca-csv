import SPListViewService from "./listService/listService";
import { format } from 'date-fns';

const generateCSVdata = async (viewItemIds: string[], templateData: any[], templateType: string): Promise<any[]> => {

    const savedRows: string[] = [];

    const getSumma = (templateRow: string[], listItems: any[]) => {
        const csvName: string = templateRow[13];
        const currency: string = templateRow[8];
        let result: number = 0;
        if (savedRows.indexOf(csvName + currency + (templateRow[14].indexOf('betét') > -1 ? 'betet' : 'iroda'))) {
            listItems.forEach(item => {
                if (item.Affiliate.CSVName === csvName && item.Currency.Title === currency) {
                    if (templateRow[14].indexOf('betét') > -1) {
                        result = result + item.BankDeposit;
                    } else if (templateRow[6] === "PÉNZÁLL") {
                        result = result + (item.CurrentAccounts + item.CashOnHandAndCheques);
                    }
                }
                if (item.Affiliate.CSVName === csvName && item.Currency2.Title === currency) {
                    if (templateRow[14].indexOf('betét') > -1) {
                        result = result + item.BankDeposit2;
                    } else if (templateRow[6] === "PÉNZÁLL") {
                        result = result + (item.CurrentAccounts2 + item.CashOnHandAndCheques2);
                    }
                }
                if (item.Affiliate.CSVName === csvName && item.Currency3.Title === currency) {
                    if (templateRow[14].indexOf('betét') > -1) {
                        result = result + item.BankDeposit3;
                    } else if (templateRow[6] === "PÉNZÁLL") {
                        result = result + (item.CurrentAccounts3 + item.CashOnHandAndCheques3);
                    }
                }
                templateRow[3] = format(new Date(item.DateOfBalance), 'yyyy.MM.dd');
                templateRow[4] = format(new Date(item.DateOfBalance), 'yyyy.MM.dd');
                templateRow[5] = format(new Date(item.DateOfBalance), 'yyyy.MM.dd');
            });
            savedRows.push(csvName + currency + (templateRow[14].indexOf('betét') > -1 ? 'betet' : 'iroda'));
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
        let dateOfBalance: Date;

        listItems.forEach(item => {    
            if (item.Currency.Title === currency) {
                result = result + (item.CurrentAccounts + item.CashOnHandAndCheques + item.BankDeposit);
                dateOfBalance = item.DateOfBalance ? new Date(item.DateOfBalance) : null;
            }
            if (item.Currency2.Title === currency) {
                result = result + (item.CurrentAccounts2 + item.CashOnHandAndCheques2 + item.BankDeposit2);
                dateOfBalance =  item.DateOfBalance ? new Date(item.DateOfBalance) : null;
            }
            if (item.Currency3.Title === currency) {
                result = result + (item.CurrentAccounts3 + item.CashOnHandAndCheques3 + item.BankDeposit3);
                dateOfBalance =  item.DateOfBalance ? new Date(item.DateOfBalance) : null;
            }
        });
        templateRow[3] = dateOfBalance ? format(dateOfBalance, 'yyyy.MM.dd') : "";
        templateRow[4] = dateOfBalance ? format(dateOfBalance, 'yyyy.MM.dd') : "";
        templateRow[5] = dateOfBalance ? format(dateOfBalance, 'yyyy.MM.dd') : "";
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