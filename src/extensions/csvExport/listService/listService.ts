import { sp } from "@pnp/sp";  
import { Web } from "@pnp/sp/webs";  
import "@pnp/sp/webs";
import "@pnp/sp/lists";  
import "@pnp/sp/items";  
import "@pnp/sp/site-users/web";
  
export class listService {

    public context:any;
  
    public setup(context: any): void {  
        sp.setup({  
            spfxContext: context  
        });
        this.context = context;
    }

    public async CurrentUserGroups(): Promise<any> {
        return new Promise<any>((resolve, reject) => {
            try {
                sp.web.currentUser.groups().then((groups) => {
                    resolve(groups);
                });
            }
            catch (error) {
                console.log(error);
            }
        });
    }

    public async isGroupMember(spGroupsTitle:string[]):Promise<boolean> {
        return new Promise<boolean>((resolve, reject) => {
            try {
                if(spGroupsTitle.length === 0) {
                    resolve(true);
                } else {
                    let  isMember:boolean = false;
                    this.CurrentUserGroups().then((result:any[]) => {
                        result.forEach((group:any) => {
                            if(spGroupsTitle.indexOf(group.Title) > -1){
                                isMember = true;
                            }
                        });
                        resolve(isMember);
                    });
                }
            }
            catch (error) {
                console.log(error);
            }
        });
    }

    public async getListItems(itemIDs:string[]):Promise<any[]> {
        return new Promise<any[]>((resolve, reject) => {
            try{
                let filter = "";
                itemIDs.forEach((itemID:string, index:number) => {
                    if(index === itemIDs.length - 1){
                        filter += `ID eq ${itemID}`
                    } else {
                        filter += `ID eq ${itemID} or `
                    }
                });
                const select = '*, Affiliate/ID, Affiliate/Title, Affiliate/CSVName, Affiliate/TipusText, Currency/ID, Currency/Title';
                const expand = 'Affiliate, Currency';
                sp.web.lists.getByTitle("Weekly closing balance report").items.filter(filter).select(select).expand(expand).get().then((items) => {
                    resolve(items);
                })
            }
            catch (error) {
                console.log(error);
            }
        })
    }
    
}  
  
const SPListViewService = new listService();  
export default SPListViewService;