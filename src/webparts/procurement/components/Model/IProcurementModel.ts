export interface IProcurementModel{
  PreferredVendor:string;
  ItemDescription: string;
  SiteLink:string;
  Quantity:any;
  UnitCost:any;
  TotalCost:any;
  RequestorID:string;
  Status:string;
  Manager:string;
  Creator:string;
  ID:number;
  HighManagementStatus:string;
  ManagerStatus:string;
  ReviewForHighManagement:string;
  //AmountPaidDate:any;
}
export interface IProcurementListItems{
  Phone:string;
  BlanketPORequest: string;
  PaymentFrom:string;
  DateRequired:any;
  ShipAddress:string;
  PurchasingEntity:string;
  PayingEntity:string;
  OtherAddress:string;
  RequestorID:string;
  Status:string;
  Manager:string;
  Creator:string;
  ID:number;
  HighManagementStatus:string;
  ManagerStatus:string;
  AccountantStatus:string;
  ReviewForHighManagement:string;
  Accountant:string;
  TotalExpense:any;
  OrderedDate:any;
  //AmountPaidDate:any;
}
export interface ILogHistoryModel{
  Id:any;
  RequestID:any;
  Procurement:any;
  Author:any;
  CreatedOn:any;
  Status:string;
  CommentsHistory:any;
}