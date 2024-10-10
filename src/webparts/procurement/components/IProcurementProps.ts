import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IProcurementProps {
  //description: string;
  siteUrl:string;
  context:WebPartContext;
  procurementRequestList:string;
  procurementReqDetailList:string;
  logHistoryListTitle:string;
  
}
