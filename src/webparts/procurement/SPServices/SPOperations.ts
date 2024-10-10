import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp, Web } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IProcurementListItems, IProcurementModel } from "../components/Model/IProcurementModel";


export class SPOperations {

    public constructor(public siteUrl: string) { }
     //** Convert dates */
  public ConvertDate(dateValue){
    var d = new Date(dateValue);
    var strDate =  d.getDate()+ "/" + (d.getMonth()+1) + "/" + d.getFullYear();
    return strDate;
  }
  //* get Log History by Expense ID **/
public GetLogHistoryItems(itemId: any,listName:string):Promise<any> {
  let logHistoryItems:any[]=[];
  let web = Web(this.siteUrl);
  return new Promise<any>(async(resolve,reject)=>{
    web.lists.getByTitle(listName).items.filter(`ProcurementId eq `+itemId).select("*","Author/Title","Author/ID","Author/EMail","Procurement/ID","Procurement/Title").expand("Author,Procurement").orderBy("Id",false).get().then(results => {
      results.map((item)=>{
        logHistoryItems.push({
          Id:item.ID,
          RequestID:item.Title,
          Procurement:item.Procurement!=undefined?item.Procurement.Title:"",
          Author:item.Author.Title,
          CreatedOn:this.ConvertDateYYMMDD(item.Created),
          Status:item.Status,
          CommentsHistory:item.CommentsHistory!=undefined?item.CommentsHistory.replace(/<\/?[^>]+(>|$)/g, ""):"",
        });
      })
      resolve(logHistoryItems);
},(error:any)=>{
    reject("error occured "+error);
})
})
};
 //* get ListItems by login name **/
 public getListItems(email: string,requestType,userType):Promise<IProcurementListItems[]> {
  let listItems:IProcurementListItems[]=[];
  let query:string="";
  if(requestType=="MySubmission"){
   query=`Author/EMail eq '${email}'`;
  }
  if(requestType=="MyTask"){
    if(userType=="Manager"){
    query=`Manager/EMail eq '${email}'`;
    }
    // if(userType=="HighManagement"){
    //   query=`(HighManagement/EMail eq '${email}' and ReviewForHighManagement eq 'Yes') and ((ManagerStatus eq 'Approved') or (ManagerStatus eq 'Rejected by HighManagement') or (ManagerStatus eq 'Manager Approval Not Required'))`;
    //   }
  }
  let web = Web(this.siteUrl);
  return new Promise<any>(async(resolve,reject)=>{
  web.lists.getByTitle('ProcurementRequests').items
  .filter(query)
  .select("*","Manager/Title","Manager/ID","Manager/EMail","Author/Title","Author/ID","Author/EMail","HighManagement/Title","HighManagement/ID","HighManagement/EMail","Accountant/Title","Accountant/ID","Accountant/EMail").expand("Manager,Author,HighManagement,Accountant")
  .orderBy("Modified",false)
  .top(4999)
  .get().then(results => {
    console.log(listItems);
    if(userType=="HighManagement"){
      let allProcurementItems=results;
      results=[];
      allProcurementItems.map((ProcurementItem)=>{
        if((ProcurementItem.ReviewForHighManagement =="Yes") && (ProcurementItem.ManagerStatus=="Approved" || ProcurementItem.ManagerStatus=="Rejected by HighManagement"|| ProcurementItem.ManagerStatus=="Manager Approval Not Required")){
        if(ProcurementItem.HighManagement!=undefined){
          ProcurementItem.HighManagement.map((highMgmt)=>{
          if(highMgmt.EMail==email){
            results.push(ProcurementItem);
          }
        })
      }
      }
      })
    }
    if(userType=="Accountant"){
      let allAccountantItems=results;
      results=[];
      allAccountantItems.map((AccountantItem)=>{
        if((AccountantItem.HighManagementStatus=="Approved" || AccountantItem.Status=="Pending for Accountant" || AccountantItem.AccountantStatus=="Order Placed" || AccountantItem.AccountantStatus=="Rejected"|| AccountantItem.AccountantStatus=="Pending for Accountant")){
          if(AccountantItem.Accountant!=undefined){
          AccountantItem.Accountant.map((accountant)=>{
          if(accountant.EMail==email){
            results.push(AccountantItem);
          }
        })
      }
      }
      })
    }
    results.map((item)=>{
      let StatusValue="";
      if(requestType=="MySubmission"){
        StatusValue=item.Status;
       }
      if(requestType=="MyTask"){
        if(userType=="Manager"){
          StatusValue=item.ManagerStatus;
        }
        if(userType=="HighManagement"){
          StatusValue=item.HighManagementStatus;
          }
          if(userType=="Accountant"){
            StatusValue=item.AccountantStatus;
          }
      }
      listItems.push({
        Phone:item.Phone,
        BlanketPORequest: item.Title,
        TotalExpense:item.TotalExpense,
        PaymentFrom:item.PaymentFrom,
        PurchasingEntity:item.PurchasingEntity,
        PayingEntity:item.PayingEntity,
        DateRequired:item.DateRequired!=null?this.ConvertDate(item.DateRequired):null,
        ShipAddress:item.ShipAddress,
        OtherAddress:item.OtherAddress,
        RequestorID:item.RequestorID,
        Status:StatusValue,
        Manager:item.Manager!=undefined?item.Manager.Title:"",
        Creator:item.Author!=undefined?item.Author.Title:"",
        ID:item.ID,
        HighManagementStatus:item.HighManagementStatus,
        ReviewForHighManagement:item.ReviewForHighManagement,
        ManagerStatus:item.ManagerStatus,
        AccountantStatus:item.AccountantStatus,
        Accountant:item.Accountant,
        OrderedDate:item.OrderedDate!=null?this.ConvertDate(item.OrderedDate):null,
      });
    })
    resolve(listItems)
},(error:any)=>{
    reject("error occured "+error);
})
})
};

//* get ListItems by login name **/
public _getProcurementConfigItems():Promise<any[]> {
  let web = Web(this.siteUrl);
  return new Promise<any>(async(resolve,reject)=>{
  web.lists.getByTitle('ProcurementConfigurations').items
  .select("*","HighManagement/Title","HighManagement/ID","HighManagement/EMail","EmailIds/Title","EmailIds/ID","EmailIds/EMail").expand("HighManagement,EmailIds")
  .orderBy("Modified",false)
  .top(4999)
  .get().then(results => {
    resolve(results)
},(error:any)=>{
    reject("error occured "+error);
})
})
};
  /**
     * CreateProcurement Item
     */
   public async CreateItem(listTitle:string,data:any): Promise<any> {
    let web = Web(this.siteUrl);
    return new Promise<string>(async (resolve, reject) => {
      await web.lists.getByTitle(listTitle).items.add(data)
        .then((result: any) => {
          resolve(result)
        }, (error: any) => {
          reject("error occured " + error);
        })
    })
  };
   //** Convert dates */
 public ConvertDateYYMMDD(dateValue){
  var d = new Date(dateValue),
     month = '' + (d.getMonth() + 1),
     day = '' + d.getDate(),
     year = d.getFullYear();

 if (month.length < 2) month = '0' + month;
 if (day.length < 2) day = '0' + day;

 return [year, month, day].join('-');
};
  //* get Procurement detail by lookup ID **/
public GetProcurementDetails(itemId: any,listName:string):Promise<any> {
  let listItems:any[]=[];
  let web = Web(this.siteUrl);
  return new Promise<any>(async(resolve,reject)=>{
    web.lists.getByTitle(listName).items.select("*").filter("ProcurementID eq "+itemId+"").get().then(results => {
      results.map((item)=>{
        listItems.push({
          id:item.Id,
          Id:item.Id,
          PreferredVendor:item.PreferredVendor,
          ItemDescription: item.Title,
          SiteLink:item.SiteLink,
          Quantity:item.Quantity,
          UnitCost:item.UnitCost,
          TotalCost:item.TotalCost,
          Status:item.Status,
        });
      })
      resolve(listItems);
},(error:any)=>{
    reject("error occured "+error);
})
})
};
//* get ListItem by Item ID **/
public GetListItemByID(itemId: any,listName:string):Promise<any> {
  let web = Web(this.siteUrl);
  return new Promise<any>(async(resolve,reject)=>{
    web.lists.getByTitle(listName).items.getById(itemId).select("*","Manager/Title","Manager/ID","Manager/EMail","Author/Title","Author/ID","Author/EMail","HighManagement/Title","HighManagement/ID","HighManagement/EMail","Accountant/Title","Accountant/ID","Accountant/EMail").expand("Manager,Author,HighManagement,Accountant").get().then(results => {
    console.log(results);
    resolve(results)
},(error:any)=>{
    reject("error occured "+error);
})
})
};
    /**
     * updateItem
     */
     public UpdateItem(listTitle:string,data:any,itemId:any):Promise<string> {
        let web=Web(this.siteUrl);
        return new Promise<string>(async(resolve,reject)=>{
          web.lists.getByTitle(listTitle).items.getById(itemId).update(data)
          .then((result:any)=>{
              resolve("Updated")
          },(error:any)=>{
              reject("error occured "+error);
          })
        })
    };
    //** Get Today Date */
public GetTodaysDate (){
    var d = new Date();
    let month = '' + (d.getMonth() + 1),
    day = '' + d.getDate(),
    year = d.getFullYear();
  
  if (month.length < 2) month = '0' + month;
  if (day.length < 2) day = '0' + day;
  
  return [day, month, year].join('-');
  }
     //**Get Current User**/
     public GetCurrentUser() :Promise<string> {
        let web = Web(this.siteUrl);
        return new Promise<string>(async(resolve,reject)=>{
        web.currentUser.get().then((result: any) => {
          resolve(result)
        },(error:any)=>{
            reject("error occured "+error);
        })
      })
  };
        //*get Current User Details **/
        public getCurrentUserDetails(empName:string):Promise<string>{
          let web = Web(this.siteUrl);
            return new Promise<string>(async(resolve,reject)=>{
          sp.profiles.getPropertiesFor(empName).then((profile: any) => {
          var properties = {};
          profile.UserProfileProperties.forEach(function(prop) {
          properties[prop.Key] = prop.Value;
          });
          resolve(properties["Manager"])
        },(error:any)=>{
            reject("error occured "+error);
        })
      })
  };
    //* get Manager Details**/
    public getManagerDetails(user:string):Promise<string>{
        return new Promise<string>(async(resolve,reject)=>{
        sp.profiles.getPropertiesFor(user).then((profile: any) => {
        var properties = {};
        profile.UserProfileProperties.forEach(function(prop) {
        properties[prop.Key] = prop.Value;
        });
        resolve(properties["WorkEmail"])
      },(error:any)=>{
          reject("error occured "+error);
      })
    })
};
       //* get User ID by Email **/
       public getUserIDByEmail(email: string):Promise<any> {
        let web = Web(this.siteUrl);
        return new Promise<any>(async(resolve,reject)=>{
        web.siteUsers.getByEmail(email).get().then(user => {
          console.log('User Id: ', user.Id);
          resolve(user.Id)
      },(error:any)=>{
          reject("error occured "+error);
      })
    })
};
}
