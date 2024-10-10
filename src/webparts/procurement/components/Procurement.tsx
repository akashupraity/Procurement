import * as React from 'react';
import styles from './Procurement.module.scss';
import { IProcurementProps } from './IProcurementProps';
import { IProcurementState } from './IProcurementState';
import { escape } from '@microsoft/sp-lodash-subset';

import * as $ from "jquery";
import * as bootstrap from "bootstrap";
require('../../../../node_modules/bootstrap/dist/css/bootstrap.min.css');
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IItem, sp, Web } from '@pnp/sp/presets/all'
import { SPOperations } from '../SPServices/SPOperations';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { IStyleSet, mergeStyleSets, getTheme, FontWeights, } from 'office-ui-fabric-react/lib/Styling';
import { Label, ILabelStyles } from 'office-ui-fabric-react/lib/Label';
import { IPivotItemProps, Pivot, PivotItem, PivotLinkFormat } from 'office-ui-fabric-react/lib/Pivot';
import { ListView, IViewField, SelectionMode, GroupOrder, IGrouping } from "@pnp/spfx-controls-react/lib/ListView";
import {
  DefaultButton,
  Modal,
  IconButton, IButtonStyles, IIconProps, IDetailsRowStyles, DetailsRow, Icon
} from 'office-ui-fabric-react';
import { jsPDF } from "jspdf";
import html2canvas from 'html2canvas';
import { DateTimePicker, DateConvention, TimeConvention, TimeDisplayControlType } from '@pnp/spfx-controls-react/lib/dateTimePicker'; 

const theme = getTheme();
const contentStyles = mergeStyleSets({
  container: {
    display: 'flex',
    // flexFlow: 'column nowrap',
    alignItems: 'stretch',
    width: '450px',
    //height: '260px',
    color: '#000',
    padding: '10px',
    overflow: 'hidden',
  },
  header: [
    // eslint-disable-next-line deprecation/deprecation
    theme.fonts.xLargePlus,
    {
      flex: '1 1 auto',
      //borderTop: `4px solid ${theme.palette.themePrimary}`,
      color: theme.palette.neutralPrimary,
      display: 'flex',
      alignItems: 'center',
      fontWeight: FontWeights.semibold,
      // padding: '12px 12px 14px 24px',
      fontSize: '20px',
      overflow: 'hidden',
    },
  ],
  body: {
    flex: '4 4 auto',
    padding: '0 24px 24px 24px',
    overflowY: 'hidden',
    selectors: {
      p: { margin: '14px 0' },
      'p:first-child': { marginTop: 0 },
      'p:last-child': { marginBottom: 0 },
    },
  }
})
const iconButtonEditFormStyles: Partial<IButtonStyles> = {
  root: {
    color: theme.palette.neutralPrimary,
    marginLeft: 'auto',
    marginTop: '4px',
    marginRight: '2px',
  },
  rootHovered: {
    color: theme.palette.neutralDark,
  },
};
const contentStatusBarStyles = mergeStyleSets({
  container: {
    display: 'flex',
    // flexFlow: 'column nowrap',
    alignItems: 'stretch',
    width: '1060px',
    height: '250px',
    color: '#000',
    padding: '10px',
    overflow: 'hidden',
  },
  header: [
    // eslint-disable-next-line deprecation/deprecation
    theme.fonts.xLargePlus,
    {
      flex: '1 1 auto',
      color: theme.palette.neutralPrimary,
      display: 'flex',
      alignItems: 'center',
      fontWeight: FontWeights.semibold,
      // padding: '12px 12px 14px 24px',
      fontSize: '20px',
      overflow: 'hidden',
    },
  ],
  body: {
    flex: '4 4 auto',
    padding: '0 24px 24px 24px',
    overflowY: 'hidden',
    selectors: {
      p: { margin: '14px 0' },
      'p:first-child': { marginTop: 0 },
      'p:last-child': { marginBottom: 0 },
    },
  },
  label: {
    'font-family': 'inherit',
    'font-weight': '600',
  },
  'h5':{
    'margin-left': '-12px',
    'margin-top': '15px',
    'font-size': 'small',
  },
  'h4':{
    'margin-left': '-13px',
    'font-size': '16px',
    'margin-top': '-20px',
    'color': 'blue',
    'width': 'max-content',
  },
  'h3':{
    'margin-left': '-13px',
    'font-size': '16px',
    'margin-top': '-20px',
    'color': 'blue',
    'width': '240px',
  },
  'mailIcon':{
    'padding-left': '36px',
  },
  'progressbar': {
    'width': '809px',
    'height': '9px',
    'background-color': 'lightgray',
    'border-radius': '10px',
    'display': 'flex',
    'align-items': 'center',
    'justify-content': 'space-between',
    'padding': '0px',
    'margin-top': '5%',
    'margin-left': '5%',
},
'Paid':{
  'margin-left': '18px',
},
'statuscircle': {
  'width': '40px', /* Adjust the size as per your requirement */
  'height': '40px', /* Adjust the size as per your requirement */
  'border-radius': '50%',
  'background-color': 'lightgray', /* Default color */
  'transition': 'background-color 0.3s',
},
'approved': {
  'background-color': 'green', /* Change color when approved */
},
'rejected': {
  'background-color': 'red', /* Change color when rejected */
},
'pending': {
  'background-color': '#ffb100', /* Change color when pending */
},
'clarify':{
  'background-color': '#09cceb', /* Change color when clarification */
},
'icon':{
  'margin-top': '12px',
    'margin-left': '12px',
    'width': '10px',
    'height': '10px',
    'color': 'white',
}
})

const liTheme = getTheme();
// const labelStyles: Partial<IStyleSet<ILabelStyles>> = {
//   root: { marginTop: 10 },
// };
const cancelIcon: IIconProps = { iconName: 'Cancel' };

//const theme = getTheme();
const contentEditFormStyles = mergeStyleSets({
  container: {
    display: 'flex',
    // flexFlow: 'column nowrap',
    alignItems: 'stretch',
    width: '1200px',
    height: '800px',
    color: '#000',
    padding: '10px',
  },
  header: [
    // eslint-disable-next-line deprecation/deprecation
    theme.fonts.xLargePlus,
    {
      flex: '1 1 auto',
      color: theme.palette.neutralPrimary,
      display: 'flex',
      alignItems: 'center',
      fontWeight: FontWeights.semibold,
      //  padding: '12px 12px 14px 24px',
      fontSize: '20px',
      
    },
  ],
  linkImg:{
   'width':'10px',
   //'margin-left':'-75px'
  },
  statusImg:{
    'width':'18px',
  },
  loaderEdit: {
    'display': 'none',
    'position': 'fixed',
    'top': '0',
    'left': '0',
    'right': '0',
    'bottom': '0',
    'width': '100%',
    'background': "rgba(0,0,0,0.75) url('https://bbsmidwestcom.sharepoint.com/sites/Operation/SiteAssets/ICONS/loading2.gif') no-repeat center center",
    'z-index': '10000',
  },
  link:{
    'float':'right',
    'margin-right': '-16px',
    'margin-top': '-30px',
  },
  commentContainer: {
    'margin': '0px auto',
    'background': '#f5f4f7',
    'border-radius': '8px',
    'padding': '14px',
  },
  cmtHistoryRow: {
    //'display': '-ms-flexbox',
    'display': 'flex',
    '-ms-flex-wrap': 'wrap',
    'flex-wrap': 'wrap',
    'margin-right': '-13px',
    'margin-left': '-5px',
    'overflow-y': 'auto',
    'max-height': '120px',
  },
  comment: {
    'display': 'block',
    'transition': 'all 1s',
  },
  viewInvoice:{
    'margin-bottom': '7px',
  },
  label: {
    'font-family': 'inherit',
    'font-weight': '600',
  },
  lblClr: {
    'font-size': '15px',
    'color': '#007bff',
    'font-family': 'inherit',
    'font-weight': '600',
    'width':'80px'
  },
  newRequestTable: {
    'width': '100%',
  },
  deptSection: {
    'background-color': '#f7f7f7',
    //  'display': '-ms-flexbox',
    'display': 'flex !important',
    '-ms-flex-wrap': 'wrap',
    'flex-wrap': 'wrap',
    'margin-right': '-5px',
    'margin-left': '-5px',
    'margin-bottom': '10px',
    'margin-top': '5px',
  },
  viewDeptSection: {
    'background-color': '#f7f7f7',
    'display': 'flex !important',
    'flex-wrap': 'wrap',
    'margin-right': '-5px',
    'margin-left': '-5px',
    'margin-bottom': '15px',
    'margin-top':'15px',
  },
  editRequestTable: {
    // 'display': '-ms-flexbox',
    'display': 'flex',
    '-ms-flex-wrap': 'wrap',
    'flex-wrap': 'wrap',
    'margin-right': '-5px',
    'margin-left': '-5px',
    'margin-top': '-10px',
    'background-color': '#fdfdfd',
  },
  line: {
    'width': '100%',
    'border-top': '1px solid rgba(0,0,0,.1)',
    'display': 'inline-block',
  },
  noteMsg: {
    'font-family': 'inherit',
    'font-weight': '400',
    'font-size': '9px',
    'color': '#a5a3a3',
  },
  formRow: {
    //'display': '-ms-flexbox',
    'display': 'flex',
    '-ms-flex-wrap': 'wrap',
    'flex-wrap': 'wrap',
    'margin-right': '-5px',
    'margin-left': '-5px',
    'margin-bottom': '-10px',
    'background-color': '#fdfdfd',
  },
  lblCtrl: {
    'margin-top': '5px',
    'font-family': 'inherit',
    'font-weight': '600',
    'width': '170px'
  },
  pr15: {
    'padding-right': '15px',
  },
  btnbr9: {
    'border-radius': '9px',
    'color': '#fff',
    'background-color': '#007bff',
    'border-color': '#007bff',
  },
  ml8: {
    'margin-left': '8px',
  },


  star: {
    'color': 'red',
  },
  errMsg: {
    'color': 'red',
  },
  deleteIcon: {
    'color': 'red',
    'margin-top': '10px',
  },
  btnRt: {
    float: 'left',
    'margin-right': '15px',
    'margin-bottom': '15px',
  },


  inputAttachment: {
    'padding-bottom': '25px',
  },
  expenselbl: {
    'font-weight': '600',
    'color': 'white',
    'font-size': 'initial',
    'background-color': '#6492c3',
    'padding-left': '5px',
  },
  cmt: {
    'width': '300px',
    'padding-right': '15px',
  },
  attachFile: {
    'width': '150px',
    'padding-right': '15px',
  },
  attachedFile: {
    'width': '500px',
    'padding-right': '15px',
  },
  formControl: {
    'display': 'block',
    'width': '130%',
    'height': 'calc(1.5em + .75rem + 2px)',
    'padding': '.375rem .75rem',
    'font-size': '1rem',
    'font-weight': '400',
    'line-height': '1.5',
    'color': '#495057',
    'background-color': '#fff',
    'background-clip': 'padding-box',
    'border': '1px solid #ced4da',
    'border-radius': '.25rem',
    'transition': 'border-color .15s ease-in-out,box-shadow .15s ease-in-out'
  },
  itemLine:{
    'width': '100%',
    'border-top': '1px solid rgba(0,0,0,.1)',
    'display': 'inline-block',
    'margin-bottom': '15px',
  }
});
const contentInvoiceStyles = mergeStyleSets({
  container: {
    display: 'flex',
    // flexFlow: 'column nowrap',
    alignItems: 'stretch',
    width: '1060px',
    //height: '260px',
    color: '#000',
    padding: '10px',
    overflow: 'hidden',
  },
  header: [
    // eslint-disable-next-line deprecation/deprecation
    theme.fonts.xLargePlus,
    {
      flex: '1 1 auto',
      color: theme.palette.neutralPrimary,
      display: 'flex',
      alignItems: 'center',
      fontWeight: FontWeights.semibold,
      // padding: '12px 12px 14px 24px',
      fontSize: '20px',
      overflow: 'hidden',
    },
  ],
  body: {
    flex: '4 4 auto',
    padding: '0 24px 24px 24px',
    overflowY: 'hidden',
    selectors: {
      p: { margin: '14px 0' },
      'p:first-child': { marginTop: 0 },
      'p:last-child': { marginBottom: 0 },
    },
  },
  'invoiceNum':{
    'margin-top': '60px',
    'margin-left': '60px',
  },
  'table' : {
    'border-collapse': 'collapse',
    'width': '90%',
    'margin-top': '0px',
    'border': '2px solid #787878',
    'align-items': 'center',
  },
  'table2': {
    'border-collapse': 'collapse',
    'width': '90%',
    'margin-top': '-10px',
    'border': '2px solid #787878',
  },
  'th, td' : {
    'padding': '8px',
    'text-align': 'left',
    'border-bottom': '1px solid #ddd',
    'font-weight': '100',
    'font-size': 'small',
  },
  'td' : {
    'padding': '8px',
    'text-align': 'left',
    'border-bottom': '1px solid #ddd',
    'font-weight': '100',
    'font-size': 'small',
  },
  'th': {
    'background-color': 'white',
  },
  h1 : {
    'text-align': 'center',
  },
  'invoiceinfo': {
    'margin-bottom': '20px',
  },
  'invoiceinfo p': {
    'margin': '0'
  },
  'invoiceinfo p strong' :{
    'margin-right': '10px',
  },
  'total': {
    'font-weight': 'bold',
  },
  'Container':{
    'width': '90%',
    'height': '35px',
    'color': 'white',
    'background-color': '#484848',
    'text-align': 'left',
    'border-radius': '0px',
    'font-size': '15px',
  },
  'Container2':{
    'width': '90%',
    'height': '45px',
    'color': 'white',
    'background-color': '#484848',
    'text-align': 'left',
    'border-radius': '1px',
    'font-size': '15px',
  },
  h3:{
    'padding-left': '20px',
    'padding-top': '10px',
  },
  'Containerbox': {
    'width': '100%',
    'margin': '0 auto', /* Center the container horizontally */
    'padding': '20px',
    'border': '1px solid #ccc',
    'padding-right': '-5%',
    'padding-left': '80px',
  },
  'footeritem': {
    'text-align':'right',
   'padding-right': '10%',
   'padding-bottom': '5%',
		//'margin-top': '-10%',
  }
})
const iconButtonStyles: Partial<IButtonStyles> = {
  root: {
    color: theme.palette.neutralPrimary,
    marginLeft: 'auto',
    marginTop: '4px',
    marginRight: '2px',
  },
  rootHovered: {
    color: theme.palette.neutralDark,
  },
};

export default class Procurement extends React.Component<IProcurementProps, IProcurementState> {
  public _spOps: SPOperations;
  constructor(props: IProcurementProps) {
    super(props)
    SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css');
    SPComponentLoader.loadCss("https://maxcdn.bootstrapcdn.com/bootstrap/3.4.0/css/bootstrap.min.css");
    SPComponentLoader.loadCss("//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css");
    this.state = {
      IProcurementModel: [],
      Comments: "",
      Phone: "",
      BlanketPORequest: "",
      PaymentFrom: "",
      DateRequired: null,
      ShipAddress: "",
      PreferredVendorOptions:[],
      PaymentFromOptions:[],
      ShipAddressOptions:[],
      PayingEntityOptions:[],
      PurchasingEntityOptions:[],
      ManagerApprovalCost:"",
      HighMngmntApprovalCost:"",
      HighMngmntApprovalWithVendorsCost:"",
      PurchasingEntity:"",
      PayingEntity:"",
      OtherAddress:"",
      IsBtnClicked: false,
      SelectedTabType:"",
      fileInfos: [],
      ManagerId: null,
      filePickerResult: [],
      Attachments: [],
      SelectedFiles: [],
      CurrentUserName:"",
      openDialog: false,
      openEditDialog: false,
      HighManagementIds:null,
      selectedProcurement: {},
      latestComments:"",
      Manager:"",
      Creator:"",
      IsHighManagement:false,
      IsManager:false,
      IsAccountant:false,
      Status:"",
      TotalExpense:"",
      ProcurementItemsToDelete: [],
      ManagerStatus:"",
      FilesToDelete: [],
      PhoneErrMsg:"",
      IsPhoneErr:false,
      BlanketPORequestErrMsg:"",
      IsBlanketPORequestErr:false,
      DateRequiredErrMsg:"",
      IsDateRequiredErr:false,
      PaymentFormErrMsg:"",
      IsPaymentFormErr:false,
      ShipAddressErrMsg:"",
      IsShipAddressErr:false,
      ProcurementDetailErrMsg:"",
      IsProcurementDetailErr:false,
      AttachmentErrMsg:"",
      IsAttachmentErr:false,
      AttachmentsCounts:"",
      AttachmentsVendorNotes:"",
      IsOtherAddress:false,
      MySubmissionItems: [],
      MyPendingItems: [],
      MyApprovedItems: [],
      MyRejectedItems: [],
      MyOrderedPlacedItems:[],
      OrderedDate: null,
      IsOrderedDateErr: false,
      OrderedDateErrMsg: "",
      AllAccoutantIds:[],
      openInvoiceDialog: false,
      CreatorEmail:"",
      ILogHistoryModel:[],
      MumbaiAddressDetails:"",
      IndoreAddressDetails:"",
      DelhiAddressDetails:"",
      NoidaAddressDetails:"",
      OtherAddressDetails:"",
      HighManagementStatus:"",
      AccountantStatus:"",
      openStatusBarDialog:false,
      HighManagement:"",
      Accountant:"",
      RequestorResponse:"",
      ManagerResponse:"",
      HighManagementResponse:"",
      AccountantResponse:"",
    }

    this._spOps = new SPOperations(this.props.siteUrl);
  }

  public GetProcurementConfig() {
    this._spOps._getProcurementConfigItems().then((response: any) => {
      let AllVendors = [];
      let AllPaymentForms = [];
      let AllShipAddress = [];
      let AllCFOEmail = [];
      let AllCEOEmail = [];
      let AllCOOEmail = [];
      let AllPayingEntities=[];
      let AllPuchasingEntities =[];
      let ManagerApprovalExpense=[];
      let HighMngmtApprovalExpense=[];
      let HighMngmtApprovalWithVendors=[];
      let AttachmentsVendorDetailsCounts=[];
      let AttachmentsVendorsNotes=[];
      let Address1=[];
      let Address2=[];
      let Address3=[];
      let Address4=[];
      let Address5=[];

      response.map((item) => {
       if(item.Title=="BesaMe Wellness Management, LLC"){
        Address1.push(item.AddressDetails);
       }
       if(item.Title=="BesaMe Wellness Missouri, Inc."){
        Address2.push(item.AddressDetails);
       }
       if(item.Title=="BMD Cameron LLC"){
        Address3.push(item.AddressDetails);
       }
       if(item.Title=="BMD Liberty LLC"){
        Address4.push(item.AddressDetails);
       }
       if(item.Title=="Other"){
        Address5.push(item.AddressDetails);
       }
        if (item.Title == "CFO") {
          AllCFOEmail.push(item);
        }
        if (item.Title == "CEO") {
          AllCEOEmail.push(item);
        }
        if (item.Title == "COO") {
          AllCOOEmail.push(item);
        }
        if (item.Title == "PreferredVendor") {
          let vendor = { key: "", text: "" };
          vendor.key = item.Key;
          vendor.text = item.Value
          AllVendors.push(vendor);
        }
        if (item.Title == "PaymentForm") {
          let paymentFormItems = { key: "", text: "" };
          paymentFormItems.key = item.Key;
          paymentFormItems.text = item.Value
          AllPaymentForms.push(paymentFormItems);
        }
        if (item.Title == "ShipAddress") {
          let shipAddressItems = { key: "", text: "" };
          shipAddressItems.key = item.Key;
          shipAddressItems.text = item.Value
          AllShipAddress.push(shipAddressItems);
        }
        if (item.Title == "PayingEntity") {
          let payingEntityItems = { key: "", text: "" };
          payingEntityItems.key = item.Key;
          payingEntityItems.text = item.Value
          AllPayingEntities.push(payingEntityItems);
        }
        if (item.Title == "PurchasingEntity") {
          let purchasingEntityItems = { key: "", text: "" };
          purchasingEntityItems.key = item.Key;
          purchasingEntityItems.text = item.Value
          AllPuchasingEntities.push(purchasingEntityItems);
        }
        if (item.Title == "ManagerApproval") {
          ManagerApprovalExpense.push(item.Key);
        }
        if (item.Title == "HighMngmtApproval") {
          HighMngmtApprovalExpense.push(item.Key);
        }
        if (item.Title == "HighMngmtApprovalWithVendorDetails") {
          HighMngmtApprovalWithVendors.push(item.Key);
        }
        if (item.Title == "AttachmentsVendorDetails") {
          AttachmentsVendorDetailsCounts.push(item.Key);
        }
        if (item.Title == "AttachmentsVendorDetailsNotes") {
          AttachmentsVendorsNotes.push(item.Key);
        }
      })
      let CFOUserId = AllCFOEmail[0].HighManagement["ID"];
      //let CEOUserId = AllCEOEmail[0].HighManagement["ID"];
      let COOUserId = AllCOOEmail[0].HighManagement["ID"];
      let HighManagementItemIds=[];
      HighManagementItemIds.push(CFOUserId,COOUserId);
      this.setState({
        PreferredVendorOptions: AllVendors,
        PaymentFromOptions: AllPaymentForms,
        ShipAddressOptions: AllShipAddress,
        HighManagementIds:HighManagementItemIds,
        PurchasingEntityOptions:AllPuchasingEntities,
        PayingEntityOptions:AllPayingEntities,
        ManagerApprovalCost:ManagerApprovalExpense[0],
        HighMngmntApprovalCost:HighMngmtApprovalExpense[0],
        HighMngmntApprovalWithVendorsCost:HighMngmtApprovalWithVendors[0],
        AttachmentsCounts:AttachmentsVendorDetailsCounts[0],
        AttachmentsVendorNotes:AttachmentsVendorsNotes[0],
        MumbaiAddressDetails:Address1[0],
        IndoreAddressDetails:Address2[0],
        DelhiAddressDetails:Address3[0],
        NoidaAddressDetails:Address4[0],

        OtherAddressDetails:Address5[0],

      })
    })
  }
  GetCurrentUserManagerId = () => {
    this._spOps.GetCurrentUser().then((result) => {
      this.setState({
        CurrentUserName: result["Title"]
      })
      this._spOps.getCurrentUserDetails(result["LoginName"]).then((manager) => {
        if(manager!=""){
        this._spOps.getManagerDetails(manager).then((managerEmail) => {
          this._spOps.getUserIDByEmail(managerEmail).then((managerId) => {
            this.setState({
              ManagerId: managerId,
            })
          })
        })
      }
      })
    
    })
  }
  public itemExists(item, arr) {
    let isExists = false;
    arr.map((value) => {
      if (value.ID == item.ID) {
        isExists = true;
        return false;
      }
    });
    return isExists;

  }
  public getUniqueRequests(responses) {
    let uniqueResponses = [];
    responses.map((item) => {
      if (item.ManagerStatus === "Approved" && item.HighManagementStatus === "Approved" && this.itemExists(item, uniqueResponses)) {
        return false;
      }

      uniqueResponses.push(item);
      return true;
    })
    return uniqueResponses;
  }
  public GetProcurementConfigEdit(result) {
    this._spOps._getProcurementConfigItems().then((response: any) => {
      let AllHighManagementEmails=[];
      let AllCFOEmail = [];
      let AllCEOEmail = [];
      let AllCOOEmail = [];
      let AllAccountants = [];
      let AccountantIds=[];
      response.map((item) => {
        if (item.Title == "Accountant") {
          item.EmailIds.map((data)=>{
            AllAccountants.push(data.EMail);
            AccountantIds.push(data.ID)
          })
        }
        if (item.Title == "CFO") {
          AllCFOEmail.push(item);
        }
        if (item.Title == "CEO") {
          AllCEOEmail.push(item);
        }
        if (item.Title == "COO") {
          AllCOOEmail.push(item);
        }

      })
      let CFOMail = AllCFOEmail[0].HighManagement["EMail"];
      //let CEOMail = AllCEOEmail[0].HighManagement["EMail"];
      let COOMail = AllCOOEmail[0].HighManagement["EMail"];
      AllHighManagementEmails.push(CFOMail,COOMail);
      //let financeEmail = AllFinanceEmail[0].Finance["EMail"];
      let pendingItems = [];
      let rejectedItems = [];
      let approvedItems = [];
      let orderedItems = [];
      if (AllHighManagementEmails.indexOf(result["Email"])>-1) { //if manager and finance are same person
        this._spOps.getListItems(result["Email"], "MyTask", "HighManagement").then((HighManagementResponse) => {
          this._spOps.getListItems(result["Email"], "MyTask", "Manager").then((managerResponse) => {
            let responses = [...HighManagementResponse, ...managerResponse];
            let requests = this.getUniqueRequests(responses);

            requests.map((item) => {
              if ((item.Status == "Approved")) {
                approvedItems.push(item);
              }
              if (item.Status == "Rejected") {
                rejectedItems.push(item);
              }
              if (item.Status == "Pending for Manager") {
                pendingItems.push(item);
              }
              if (item.Status == "Pending for HighManagement" && parseFloat(item.TotalExpense) > parseFloat(this.state.HighMngmntApprovalCost)) {
                pendingItems.push(item);
              }
              
            })
            this.setState({
              MyPendingItems: pendingItems,
              MyRejectedItems: rejectedItems,
              MyApprovedItems: approvedItems,
              IsHighManagement: true,
              AllAccoutantIds:AccountantIds
            })
          })
        })
        //this.getMasterDetails(result);

        // })
      }
      else {
        this._spOps.getListItems(result["Email"], "MyTask", "Manager").then((response) => {
         
          if(response.length>0){
          response.map((item) => {
            if (item.Status == "Approved") {
              approvedItems.push(item);
            }
            if (item.Status == "Rejected") {
              rejectedItems.push(item);
            }
            if (item.ManagerStatus == "Pending for Manager") {
              pendingItems.push(item);
            }
          })
          this.setState({
            MyPendingItems: pendingItems,
            MyRejectedItems: rejectedItems,
            MyApprovedItems: approvedItems,
            IsManager: response.length > 0 ? true : false,
            //AllAccoutantIds:AccountantIds
          })
        }
        this.setState({
          AllAccoutantIds:AccountantIds
        })
        })
       // this.getMasterDetails(result);
      }
      if (AllAccountants.indexOf(result["Email"])>-1) { 
        this._spOps.getListItems(result["Email"], "MyTask", "Accountant").then((response) => {
          response.map((item) => {
            if (item.Status == "Order Placed") {
              orderedItems.push(item);
            }
            if (item.AccountantStatus == "Pending for Accountant") {
              pendingItems.push(item);
            }
            if (item.AccountantStatus == "Rejected") {
              rejectedItems.push(item);
            }
          })
          this.setState({
            MyPendingItems: pendingItems,
            MyOrderedPlacedItems: orderedItems,
            MyRejectedItems: rejectedItems,
            IsAccountant: response.length > 0 ? true : false,
            AllAccoutantIds:AccountantIds
          })
        })

      }
      console.log("config " + response);
      console.log("pending Items " + this.state.MyPendingItems);
    })
  }
  componentDidMount(): void {
    this._spOps.GetCurrentUser().then((result) => {
    this.GetCurrentUserManagerId();
    this.GetProcurementConfig();
    this.GetProcurementConfigEdit(result);
    this.GetMySubmittedItems(result);
   
  })
  }
  public GetMySubmittedItems(result) {
    this._spOps.getListItems(result["Email"], "MySubmission", "Owner").then((response) => {
      this.setState({
        MySubmissionItems: response
      })
    })
  };

  _handleDateChange=(date:any)=>{
   this.setState({
    DateRequired:date
   })
  };
  /* This event will fire on change of every fields on form */
  _handleChange = (index: any) => evt => {
    try {
      var item = {
        id: evt.target.id,
        name: evt.target.name,
        value: evt.target.value
      };
      // if (item.name == "PreferredVendor") {
      //   this.setState({
      //     PreferredVendor: item.value
      //   })
      // }
      if (item.name == "Phone") {
        this.setState({
          Phone: item.value
        })
      }
      if (item.name == "BlanketPORequest") {
        this.setState({
          BlanketPORequest: item.value
        })
      }
      if (item.name == "DateRequired") {
        this.setState({
          DateRequired: item.value
        })
      }
      if (item.name == "PaymentForm") {
        this.setState({
          PaymentFrom: item.value
        })
      }
      if (item.name == "ShipAddress") {
        this.setState({
          ShipAddress: item.value,
          IsOtherAddress:true,

        })
        if(item.value=="BesaMe Wellness Management, LLC"){
          this.setState({
            OtherAddress:this.state.MumbaiAddressDetails
          })
        }
        if(item.value=="BesaMe Wellness Missouri, Inc."){
          this.setState({
            OtherAddress:this.state.IndoreAddressDetails
          })
        }
        if(item.value=="BMD Cameron LLC"){
          this.setState({
            OtherAddress:this.state.DelhiAddressDetails
          })
        }
        if(item.value=="BMD Liberty LLC"){
          this.setState({
            OtherAddress:this.state.NoidaAddressDetails
          })
        }
        if(item.value=="Other"){
          this.setState({
            OtherAddress:this.state.OtherAddressDetails
          })
        }
        
        // else{
        //   this.setState({
        //     IsOtherAddress:false,
        //     OtherAddress: ""
        //   })
        // }
     
      }
      if(item.name=="OtherAddress"){
        this.setState({
          OtherAddress: item.value
        })
      }
      if (item.name == "PurchasingEntity") {
        this.setState({
          PurchasingEntity: item.value
        })
      }
      if (item.name == "PayingEntity") {
        this.setState({
          PayingEntity: item.value
        })
      }
      if (item.name == "Comments") {
        this.setState({
          Comments: item.value
        })
      }
      if (item.name == "LatestComments") {
        this.setState({
          latestComments: item.value
        })
      }
      if (item.name == "OrderedDate") {
        this.setState({
          OrderedDate: item.value
        })
      }
      var rowsArray = this.state.IProcurementModel;
      var newRow = rowsArray.map((row, i) => {
        for (var key in row) {
          if (key == item.name && row.id == item.id) {
            row[key] = item.value;
          }
          if (item.name == "UnitCost" || item.name=="Quantity") {
           // let mileDiff = parseFloat(row.EndMileout) - parseFloat(row.StartMile);
            let milageAmt = row.Quantity * row.UnitCost;
            row['TotalCost'] = milageAmt.toFixed(2);

          }
        }
        return row;
      });
      this.setState({ IProcurementModel: newRow });

    } catch (error) {
      console.log("Error in React Table handle change : " + error);
    }
  };
   //** Open Order Placed PopUp to select Order Placed date*/
   OpenOrderPlacedPopUp = () => {
    this.setState({
      openDialog: true
    })
  }
  //** Validate Order Placed Date and save it */
  CheckOrderedDate = () => {
    if (this.state.OrderedDate == null || this.state.OrderedDate == "" || this.state.OrderedDate == undefined) {
      this.setState({
        OrderedDateErrMsg: "Select Order Placed Date",
        IsOrderedDateErr: true
      })
    }
    else {
      this.UpdateRequest("Order Placed")
    }
  }
   //**Common function to render DropDowns */
   _renderDropdown = (options) => {
    return options.map((item, idx) => {
      return (<option value={item.key}>{item.text}</option>)
    })
  };
   
    // Refresh page
    _refreshNewRequestPage = () => {
      window.location.reload();
    };

  /* This event will fire on adding new row */
  _handleAddRow = () => {
    try {
      var id = (+ new Date() + Math.floor(Math.random() * 999999)).toString(36);
      const tableColProps = {
        id: id,
        PreferredVendor: "",
        ItemDescription: "",
        SiteLink:"",
        Quantity:null,
        UnitCost:null,
        TotalCost:null,

      }
      this.state.IProcurementModel.push(tableColProps);
      this.setState(this.state.IProcurementModel);
    } catch (error) {
      console.log("Error in React Table handle Add Row : " + error)
    }
  };
  /* This event will fire on remove specific row */
  _handleRemoveSpecificRow = (idx) => () => {
    try {
      const rows = this.state.IProcurementModel
      if (rows.length > 1) {
        rows.splice(idx, 1);
      }

      this.setState({ IProcurementModel: rows });
    } catch (error) {
      console.log("Error in React Table handle Remove Specific Row : " + error);
    }
  }
  private _onPivotItemClick = (item: PivotItem) => {
    if (item.props.itemKey) {
      this.setState({
        SelectedTabType: item.props.itemKey
      })
      //this.componentDidMount();
    }
  };
  public _selectedTabType(tabType: string) {
    this.setState({
      SelectedTabType: tabType
    })
  };
  // * On Select of File. Read filename and content*/
  private addFile(event) {
    //let resultFile = document.getElementById('file');
    let resultFile = event.target.files;
    console.log(resultFile);
    //let fileInfos = [];
    let fileInformations = [];
    let selectedFileNames = [];
    for (var i = 0; i < resultFile.length; i++) {
      var fileName = resultFile[i].name;
      selectedFileNames.push(fileName);
      console.log(fileName);
      var file = resultFile[i];
      var reader = new FileReader();
      reader.onload = (function (file) {
        return function (e) {
          //Push the converted file into array
          fileInformations.push({
            "name": file.name,
            "content": e.target.result
          });

        }

      })(file);

      reader.readAsArrayBuffer(file);
    }
    setTimeout(
      function () {
        let tempArr = this.state.fileInfos;
        tempArr.push.apply(tempArr, fileInformations)
        this.setState({ fileInfos: tempArr, Attachments: resultFile });
      }
        .bind(this),
      100
    );
  };

  //**Remove specific attachement */
  removeSpecificAttachment = (idx) => () => {
    try {
      const rows = this.state.fileInfos;
      rows.splice(idx, 1);
      this.setState({ fileInfos: rows });
    } catch (error) {
      console.log("Error in React Table handle Remove Specific Row : " + error);
    }
  }
  //**Remove specific attachement */
  removeSpecificAttachmentEdit = (idx, item) => () => {
    if (confirm('Are you sure, you want to remove it ?')) {
      try {
        const rows = this.state.fileInfos;
        let files = [];
        this.state.fileInfos.map((item, index) => {
          if (item.ServerRelativeUrl != undefined && index === idx) {
            files.push(item.FileName);
          }
        })
        rows.splice(idx, 1);
        let tempFileToDeleteArr = this.state.FilesToDelete;
        tempFileToDeleteArr.push.apply(tempFileToDeleteArr, files);
        console.log("Files to Remove: " + this.state.FilesToDelete);
        this.setState({ fileInfos: rows, FilesToDelete: tempFileToDeleteArr });
      } catch (error) {
        console.log("Error in React Table handle Remove Specific Row : " + error);
      }
    }
  }
  //**render selected attachments */
  renderAttachmentName() {
    return this.state.fileInfos.map((item, idx) => {
      return (<div><a href="javascript:void(0)" target="_blank">{item.name}</a>&nbsp;&nbsp; 
      <span id="delete-spec-row" onClick={this.removeSpecificAttachment(idx)} className={styles.deleteIcon}>
        <Icon iconName="delete" className="ms-IconExample" />
      </span></div>)
    })
  }
   //**render selected attachments */
   renderAttachmentNameEdit() {
    let isShowAttachDelete = this.state.SelectedTabType != "MyTask" && this.state.selectedProcurement.formType == "Edit" ? true : false;
    return this.state.fileInfos.map((item, idx) => {
      return (<div><a href={item.ServerRelativeUrl != undefined ? item.ServerRelativeUrl : "javascript:void(0)"} target="_blank" data-interception="off">{item.name}</a>&nbsp;&nbsp;
        <span id="delete-spec-row" className={contentEditFormStyles.deleteIcon}>
          {isShowAttachDelete &&
            <Icon iconName="delete" className="ms-IconExample" onClick={this.removeSpecificAttachmentEdit(idx, item)} />
          }
        </span></div>)
    })
  }
    //** Generate Requestor unique ID */
    public _getUniqueRequestorID = (procurementItemId: number) => {
      let expItemId = procurementItemId.toString();
      var uniqueID = "";
      if (procurementItemId < 10) {
        uniqueID = "000" + expItemId
      }
      if (procurementItemId >= 10 && procurementItemId < 100) {
        uniqueID = "00" + expItemId
      }
      if (procurementItemId >= 100 && procurementItemId < 1000) {
        uniqueID = "0" + expItemId
      }
      if (procurementItemId >= 1000) {
        uniqueID = expItemId;
      }
      return "PRO-" + uniqueID;
    }
     // save all procurements into "ProcurementDetails" list in batches
  private async _addProcurementDetailsAsBatch(procurements: any[], procurementItemId: number) {
    let requestorUniqueID = this._getUniqueRequestorID(procurementItemId);
    let sourceWeb = Web(this.props.siteUrl);
    let taskList = sourceWeb.lists.getByTitle(this.props.procurementReqDetailList);
    let batch = sourceWeb.createBatch();
    console.log("batch = ", JSON.stringify(batch));
    console.log("batch baseURL = ", batch["baseUrl"]);
    for (let i = 0; i < procurements.length; i++) {
      taskList.items.inBatch(batch).add(
          {
            PreferredVendor: procurements[i].PreferredVendor,
            Title: procurements[i].ItemDescription,
            SiteLink: procurements[i].SiteLink,
            Quantity: procurements[i].Quantity,
            UnitCost: procurements[i].UnitCost,
            TotalCost: procurements[i].TotalCost,
            RequestorID: requestorUniqueID,
            ProcurementIDId: procurementItemId
          }
        )
        .then((result:any) => {
          console.log("Item created with id", result.data.Id);
        })
        .catch((ex) => {
          console.log(ex);
        });
    }
    await batch.execute();
    console.log("Done");
    }
  // async function with await to save all procurements into "ProcurementDetails" list
  // private async _addProcurementDetails(procurements: any[], procurementItemId: number) {
  //   let web = Web(this.props.siteUrl);
  //   let requestorUniqueID = this._getUniqueRequestorID(procurementItemId);
  //   for (const procurement of procurements) {
  //     await web.lists.getByTitle(this.props.procurementReqDetailList).items.add({
  //       PreferredVendor: procurement.PreferredVendor,
  //       Title: procurement.ItemDescription,
  //       SiteLink: procurement.SiteLink,
  //       Quantity: procurement.Quantity,
  //       UnitCost: procurement.UnitCost,
  //       TotalCost: procurement.TotalCost,
  //       RequestorID: requestorUniqueID,
  //       ProcurementIDId: procurementItemId
  //     });
  //   }
  // }
  //* Update Unqique Id and Add Procurement details in ProcurementDetails list*/
  _updateUniqueID = (requestorUniqueID, itemId, submissionType: string) => {
    let updatePostDate = {
      RequestorID: requestorUniqueID,
    }
    this._spOps.UpdateItem(this.props.procurementRequestList, updatePostDate, itemId).then((response) => {
      if (this.state.IProcurementModel.length > 0) {
        this._addProcurementDetailsAsBatch(this.state.IProcurementModel, itemId).then(() => {
          alert(submissionType == "Submitted" ? "Request submitted sucessfully" : "Request drafted sucessfully");
          $('#loader').hide();
          this._refreshNewRequestPage();
        });
      } else {
        alert(submissionType == "Submitted" ? "Request submitted sucessfully" : "Request drafted sucessfully");
        $('#loader').hide();
        this._refreshNewRequestPage();
      }
    })
  }
    //**Validation on fields */
    ValidateForm = (submissionType) => {
      var tableArr = this.state.IProcurementModel;
      let totalAmount: any = 0;
      totalAmount=this._calculateTotalCost();
      var isErrExists = false;
      let { fileInfos } = this.state;
      this.setState({ 
        IsPhoneErr: false, 
        IsBlanketPORequestErr: false, 
        IsDateRequiredErr: false,
        IsPaymentFormErr: false, 
        IsShipAddressErr: false,
        IsProcurementDetailErr:false,
        IsAttachmentErr:false,
      });
  
      if (this.state.Phone == null || this.state.Phone == "") {
        this.setState({
          PhoneErrMsg: "Please enter Phone number",
          IsPhoneErr: true
        })
        isErrExists = true;
      }
      if (submissionType == "Submitted") {
        if (this.state.BlanketPORequest == null || this.state.BlanketPORequest == "") {
          this.setState({
            BlanketPORequestErrMsg: "Please enter Blanket PO Request",
            IsBlanketPORequestErr: true
          })
          isErrExists = true;
        }
        if (this.state.PaymentFrom == null || this.state.PaymentFrom =="") {
          this.setState({
            PaymentFormErrMsg: "Please select Payment Form",
            IsPaymentFormErr: true
          })
          isErrExists = true;
        }
        if (this.state.ShipAddress == null || this.state.ShipAddress =="") {
          this.setState({
            ShipAddressErrMsg: "Please select Ship Address",
            IsShipAddressErr: true
          })
          isErrExists = true;
        }
        if (this.state.DateRequired == null) {
            this.setState({
              DateRequiredErrMsg: "Please enter Date",
              IsDateRequiredErr: true
            })
            isErrExists = true;
        }
        if (this.state.IProcurementModel.length == 0) {
          this.setState({
            ProcurementDetailErrMsg: "Please Add Procurement Details",
            IsProcurementDetailErr: true
          })
          isErrExists = true;
        }
        
        // if(fileInfos.length == 0 && parseFloat(totalAmount)> parseFloat(this.state.HighMngmntApprovalCost)){
        //   this.setState({
        //     AttachmentErrMsg: "Please Add Attachment",
        //     IsAttachmentErr: true
        //   })
        //   isErrExists = true;
        // }
        if(fileInfos.length < parseInt(this.state.AttachmentsCounts) && parseFloat(totalAmount)> parseFloat(this.state.HighMngmntApprovalWithVendorsCost)){
          this.setState({
            AttachmentErrMsg: "Please Add Attachment with 3 vendor details",
            IsAttachmentErr: true
          })
          isErrExists = true;
        }
        // let {fileInfos}=this.state;
        tableArr.map((item, key) => {
          item.isPreferredVendorError = false;
          item.isItemDescriptionError = false;
          item.isQuantityError=false;
          item.isUnitCostError=false;
          item.isSiteLinkError = false;
          if (item.PreferredVendor == "" || item.PreferredVendor == null) {
            item.isPreferredVendorError = true;
            item.PreferredVendorErrMsg = "Please select PreferredVendor";
            isErrExists = true;
          }
          if (item.ItemDescription == "" || item.ItemDescription == null) {
            item.isItemDescriptionError = true;
            item.ItemDescriptionErrMsg = "Please enter Item Description";
            isErrExists = true;
          }
          if (item.SiteLink == "" || item.SiteLink == null) {
            item.isSiteLinkError = true;
            item.SiteLinkErrMsg = "Please enter Site Link";
            isErrExists = true;
          }
          if (item.Quantity == "" || item.Quantity == null) {
            item.isQuantityError = true;
            item.QuantityErrMsg = "Please enter Quantity";
            isErrExists = true;
          }
          if (item.UnitCost == "" || item.UnitCost == null) {
            item.isUnitCostError = true;
            item.UnitCostErrMsg = "Please enter Unit Cost";
            isErrExists = true;
          }

          // if (item.Expense == "Meal" && item.ExpenseCost >= this.state.MealExpense && this.state.fileInfos.length == 0) {
          //   this.setState({
          //     isMealExpenseCostError: true,
          //     mealExpenseCostErrMsg: "Please attach attachement"
          //   })
          //   isErrExists = true;
          // }
        })
      }
      this.setState({ IProcurementModel: tableArr });
      return isErrExists
    }
    // calculate total amount
    _calculateTotalCost=()=>{
      let totalCost: any = 0;
      if (this.state.IProcurementModel.length > 0) {
        this.state.IProcurementModel.map((amt) => {
          totalCost += amt.TotalCost != undefined ? parseFloat(amt.TotalCost) : 0;
        })
      }
      return totalCost.toFixed(2);
    }
   //** call validation method and create item into list */
   _submitRequest = (submissionType) => {
    var isError = this.ValidateForm(submissionType)
    if (!isError) {
      this.setState({
        IsBtnClicked: true
      })
      $('#loader').show();
      let commentsHTML = "";
      let todayDate = this._spOps.GetTodaysDate();
      commentsHTML = '<strong>' + this.state.CurrentUserName + ' : ' + todayDate + '</strong>' + '<div>' + this.state.Comments + '</div>';
     
      let totalAmount: any = 0;
      totalAmount=this._calculateTotalCost();
    
      let createPostData: any = {};
       createPostData = {
        Phone: this.state.Phone,
        PaymentFrom: this.state.PaymentFrom,
        Title: this.state.BlanketPORequest,
        ShipAddress:this.state.ShipAddress,
        OtherAddress:this.state.OtherAddress,
        PurchasingEntity:this.state.PurchasingEntity,
        PayingEntity:this.state.PayingEntity,
        DateRequired: this.state.DateRequired != "" ? this.state.DateRequired : null,
        //Status: submissionType == "Submitted" && this.state.ManagerId==null?"InProgress":submissionType,
       // ManagerId: this.state.ManagerId,
        //HighManagementId: {results:this.state.HighManagementIds},
        Comments: this.state.Comments != "" ? commentsHTML : this.state.Comments,
       // ManagerStatus: submissionType == "Submitted" && this.state.ManagerId!=null? "Pending for Manager" : "",
        TotalExpense: totalAmount.toString(),
      }
      // if(submissionType == "Submitted" && this.state.ManagerId==null){
      //   createPostData.HighManagementStatus="Pending for HighManagement";
      //   createPostData.ReviewForHighManagement="Yes";
      //   createPostData.ManagerStatus= 'Manager Approval Not Required';
      // }
      if(submissionType == "Submitted"){
      //if manager is not null
      if(this.state.ManagerId!=null){
        createPostData.ManagerStatus = this.state.ManagerId!=null?"Pending for Manager":"Manager Approval Not Required";
        createPostData.Status = this.state.ManagerId!=null?"Submitted":"InProgress";
        createPostData.HighManagementStatus = this.state.ManagerId!=null?"":"Pending for HighManagement";
        createPostData.ReviewForHighManagement = this.state.ManagerId!=null?"":"Yes";
        createPostData.ManagerId = this.state.ManagerId;
         }
        //if manager is null and amount > 2500
         if(totalAmount > parseInt(this.state.ManagerApprovalCost) && this.state.ManagerId==null){
          createPostData.ManagerStatus = "Manager Approval Not Required";
          createPostData.Status = "InProgress";
          createPostData.HighManagementStatus = "Pending for HighManagement";
          createPostData.ReviewForHighManagement = "Yes";
          createPostData.HighManagementId= {"results":this.state.HighManagementIds};
         }
         //if manager is null and amount < 2500
         if(totalAmount < parseInt(this.state.ManagerApprovalCost) && this.state.ManagerId==null){
          createPostData.ManagerStatus = "Manager Approval Not Required";
          createPostData.Status = "Pending for Accountant";
          createPostData.HighManagementStatus = "HighManagement Approval Not Required";
          createPostData.ReviewForHighManagement = "No";
          createPostData.AccountantStatus="Pending for Accountant";
          createPostData.AccountantId = { "results": this.state.AllAccoutantIds };
         }
        }
        if (submissionType == "Draft") {
          createPostData.ManagerStatus = "";
          createPostData.Status = "Draft";
          createPostData.ManagerId = this.state.ManagerId;
        }
        
      this._spOps.CreateItem(this.props.procurementRequestList, createPostData).then((result: any) => {
        console.log(result.data.ID);
        let itemId = result.data.ID;
        let requestorUniqueID = this._getUniqueRequestorID(itemId);
        let logHistoryPostData:any={};
        logHistoryPostData={
          Title:requestorUniqueID,
          ProcurementId:itemId,
          CommentsHistory:this.state.Comments,
          Status:createPostData.Status,
          //NameId:this.state.CurrentUserID
        }
        this._spOps.CreateItem(this.props.logHistoryListTitle, logHistoryPostData).then((result: any) => {});
        let { fileInfos } = this.state;
        if (fileInfos.length > 0) {
          let web = Web(this.props.siteUrl);
          web.lists.getByTitle(this.props.procurementRequestList).items.getById(itemId).attachmentFiles.addMultiple(fileInfos).then(() => {
            this._updateUniqueID(requestorUniqueID, itemId, submissionType);
          });
        }
        else {
          this._updateUniqueID(requestorUniqueID, itemId, submissionType);
        }
      })

    }
  }
     // update all Procurement Details in batch
     private async UpdateProcurementDetailsAsBatch(Procurements: any[]) {
      let sourceWeb = Web(this.props.siteUrl);
      let taskList = sourceWeb.lists.getByTitle(this.props.procurementReqDetailList);
      let batch = sourceWeb.createBatch();
      console.log("batch = ", JSON.stringify(batch));
      console.log("batch baseURL = ", batch["baseUrl"]);
      for (let i = 0; i < Procurements.length; i++) {
        taskList.items .getById(Procurements[i].Id).inBatch(batch).update(
            {
              PreferredVendor: Procurements[i].PreferredVendor,
              Title: Procurements[i].ItemDescription,
              SiteLink: Procurements[i].SiteLink,
              Quantity: Procurements[i].Quantity,
              UnitCost: Procurements[i].UnitCost,
              TotalCost: Procurements[i].TotalCost,
            }
          )
          .then((result:any) => {
            console.log("Item updated with id", Procurements[i].Id);
          })
          .catch((ex) => {
            console.log(ex);
          });
      }
      await batch.execute();
      console.log("Done");
      }
   // async function with await to update all Procurement Details
  //  private async UpdateProcurementDetails(Procurements: any[]) {
  //   let web = Web(this.props.siteUrl);
  //   for (const procurement of Procurements) {
  //     await web.lists.getByTitle(this.props.procurementReqDetailList).items.getById(procurement.Id).update({       
  //       PreferredVendor: procurement.PreferredVendor,
  //       Title: procurement.ItemDescription,
  //       SiteLink: procurement.SiteLink,
  //       Quantity: procurement.Quantity,
  //       UnitCost: procurement.UnitCost,
  //       TotalCost: procurement.TotalCost,

  //     });
  //   }
  // }
  public OpenStatusBarPopUp(selectedItem){
    setTimeout(
    function () {
      this.setState({ openStatusBarDialog: true, selectedProcurement: selectedItem })
      const statusCircles = document.querySelectorAll('#statuscircleId');
      //this.updateStatusCircles(statusCircles[0].childNodes,2);
      this.getSelectedProcurementDetail(selectedItem);
    }
      .bind(this),
    100
  );
}
  OpenViewFrom = (selectedItem) => {
    setTimeout(
      function () {
        selectedItem.formType = "View";
        this.setState({ openEditDialog: true, selectedProcurement: selectedItem })
        this.getSelectedProcurementDetail(selectedItem);
      }
        .bind(this),
      100
    );
    
  }
  private isEditModalOpen(): boolean {
    return this.state.openEditDialog;
  };
  private hideModal = () => {
    this.setState({
      openDialog: false
    })
  };
  private hideEditModal = () => {
    this.setState({
      openEditDialog: false
    })
    this.RefreshPage();
  }
  public selectedTabType(tabType: string) {
    this.setState({
      SelectedTabType: tabType
    })
  };
  private isInvoiceModalOpen(): boolean {
    return this.state.openInvoiceDialog;
  };
    //** Open Invoice PopUp*/
    OpenInvoicePopUp = () => {
      this.setState({
        openInvoiceDialog: true
      })
    }
     //hide Invoice popup
  private hideInvoiceModal = () => {
    this.setState({
      openInvoiceDialog: false
    })
  };
  public OpenEditForm(formType: string, selectedItem) {
    if (formType == "EditMySubmission") {
      selectedItem.formType = "Edit";
    }
    else {
      selectedItem.formType = "View";
    }
    setTimeout(
      function () {
        this.setState({ openEditDialog: true, selectedProcurement: selectedItem })
        this.getSelectedProcurementDetail(selectedItem);
      }
        .bind(this),
      100
    );
  };
  /** Render color with rows */
  private OnListViewRenderRow(props: any) {
    const customStyles: Partial<IDetailsRowStyles> = {};
    if (props) {
      if (props.itemIndex % 2 === 0) {
        //Every other row render with different background
        customStyles.root = { backgroundColor: liTheme.palette.themeLighterAlt};
      }
      return <DetailsRow {...props} styles={customStyles} />
    }
    return null;
  };

  // private onPivotItemClick = (item: PivotItem) => {
  //   if (item.props.itemKey) {
  //     this.setState({
  //       SelectedTabType: item.props.itemKey
  //     })
  //     this.componentDidMount();
  //   }
  // };
    //** Convert dates */
    public ConvertDate(dateValue) {
      var d = new Date(dateValue),
        month = '' + (d.getMonth() + 1),
        day = '' + d.getDate(),
        year = d.getFullYear();
  
      if (month.length < 2) month = '0' + month;
      if (day.length < 2) day = '0' + day;
  
      return [year, month, day].join('-');
    };
    //**get Selected Expense Item Detail */
    getSelectedProcurementDetail = (selectedItem) => {
      this._spOps.GetListItemByID(selectedItem.ID, this.props.procurementRequestList).then((result) => {
        this._spOps.GetProcurementDetails(selectedItem.ID, this.props.procurementReqDetailList).then((procurementDetails) => {
          // let AllHighManagementTitles:string="";
          // let AllAccountantTitles:string="";
          //  if(result.Accountant!=undefined){
          //   result.Accountant.map((account)=>{
          //     AllAccountantTitles+=account.Title +"; ";
          //   })
          //  }
          //  if(result.HighManagement!=undefined){
          //   result.HighManagement.map((management)=>{  
          //     AllHighManagementTitles+=management.Title +"; "
          //   })
          //  }
          this.setState({
            Phone:result.Phone,
            BlanketPORequest: result.Title,
            PaymentFrom:result.PaymentFrom,
            ShipAddress:result.ShipAddress,
            PurchasingEntity:result.PurchasingEntity,
            PayingEntity:result.PayingEntity,
            OtherAddress:result.OtherAddress,
            IsOtherAddress:true,//result.ShipAddress=="Other"?true:false,
            DateRequired: result.DateRequired != null ? this.ConvertDate(result.DateRequired) : null,
            Manager:result.Manager!=undefined ?result.Manager.Title:"",
            HighManagement:result.HighManagement!=undefined ?result.HighManagement[0].Title:"",//AllHighManagementTitles,
            Accountant:result.Accountant!=undefined ?result.Accountant[0].Title:"",//AllAccountantTitles,
            Creator:result.Author!=undefined?result.Author.Title:"",
            CreatorEmail:result.Author != undefined ? result.Author.EMail : "",
            Comments: result.Comments,
            Status: result.Status,
            TotalExpense:result.TotalExpense,
            ManagerStatus:result.ManagerStatus,
            HighManagementStatus:result.HighManagementStatus,
            AccountantStatus:result.AccountantStatus,
            IProcurementModel: procurementDetails,
            RequestorResponse:result.RequestorResponse,
            ManagerResponse:result.ManagerResponse,
            HighManagementResponse:result.HighManagementResponse,
            AccountantResponse:result.AccountantResponse,
  
          })
        })
      })
      this.getAttachments();
      this._spOps.GetLogHistoryItems(selectedItem.ID, this.props.logHistoryListTitle).then((logHistory) => {
        this.setState({
         ILogHistoryModel:logHistory
        })
       })
    };
     //**Get all attachement of Item */
  public getAttachments = () => {
    (async () => {
      // get list item by id
      const item: IItem = sp.web.lists.getByTitle(this.props.procurementRequestList).items.getById(this.state.selectedProcurement.ID);
      // get all attachments
      const attachments: any[] = await item.attachmentFiles();
      console.table(attachments);
      attachments.map((file) => {
        file.name = file.FileName
      })
      this.setState({
        fileInfos: attachments
      })
    })().catch(console.log)
  };

   // Generate PDF
   public documentprint = (e) => {  
    e.preventDefault();  
    const myinput = document.getElementById('generatePdf');  
   // html2canvas(myinput)  
     // .then((canvas) => {  
        // var imgWidth = 200;  
        // var pageHeight = 290;  
        // var imgHeight = canvas.height * imgWidth / canvas.width;  
        // var heightLeft = imgHeight;  
        // const imgData = canvas.toDataURL('image/png');  
        // const mynewpdf = new jsPDF('p', 'mm', 'a4');  
        // var position = 0;  
        // mynewpdf.addImage(imgData, 'JPEG', 5, position, imgWidth, imgHeight);  
        // mynewpdf.save("SubmittedRecord_"+this.state.selectedProcurement.RequestorID+".pdf");  
    //  });  
    html2canvas(myinput, { useCORS: true, allowTaint: true, scrollY: 0 }).then((canvas) => {
      const image = { type: 'jpeg', quality: 0.98 };
      const margin = [0.5, 0.5];
      const filename = 'myfile.pdf';

      var imgWidth = 8.5;
      var pageHeight = 11;

      var innerPageWidth = imgWidth - margin[0] * 2;
      var innerPageHeight = pageHeight - margin[1] * 2;

      // Calculate the number of pages.
      var pxFullHeight = canvas.height;
      var pxPageHeight = Math.floor(canvas.width * (pageHeight / imgWidth));
      var nPages = Math.ceil(pxFullHeight / pxPageHeight);

      // Define pageHeight separately so it can be trimmed on the final page.
      var pageHeight = innerPageHeight;

      // Create a one-page canvas to split up the full image.
      var pageCanvas = document.createElement('canvas');
      var pageCtx = pageCanvas.getContext('2d');
      pageCanvas.width = canvas.width;
      pageCanvas.height = pxPageHeight;

      // Initialize the PDF.
      var pdf = new jsPDF('p', 'in', [8.5, 11]);

      for (var page = 0; page < nPages; page++) {
        // Trim the final page to reduce file size.
        if (page === nPages - 1 && pxFullHeight % pxPageHeight !== 0) {
          pageCanvas.height = pxFullHeight % pxPageHeight;
          pageHeight = (pageCanvas.height * innerPageWidth) / pageCanvas.width;
        }

        // Display the page.
        var w = pageCanvas.width;
        var h = pageCanvas.height;
        pageCtx.fillStyle = 'white';
        pageCtx.fillRect(0, 0, w, h);
        pageCtx.drawImage(canvas, 0, page * pxPageHeight, w, h, 0, 0, w, h);

        // Add the page to the PDF.
        if (page > 0) pdf.addPage();
        debugger;
        var imgData = pageCanvas.toDataURL('image/' + image.type, image.quality);
        pdf.addImage(imgData, image.type, margin[1], margin[0], innerPageWidth, pageHeight);
      }

      pdf.save("SubmittedRecord_"+this.state.selectedProcurement.RequestorID+".pdf");
    }); 
   }
  public viewFields() {
    const viewFields: IViewField[] = [   {
      name: this.state.SelectedTabType == "MySubmission" ? "Edit" : this.state.SelectedTabType == "MyTask" ? "Action" : "",
      displayName: "",
      minWidth: 45,
      maxWidth: 45,
      render: (item: any) => {
        let isEditBtnDisable = false;
        let button;
        if (this.state.SelectedTabType == "MySubmission") {
          isEditBtnDisable = item.Status == "Rejected by HighManagement" || item.Status == "Rejected by Accountant" || item.Status == "Rejected" || item.Status == "Draft" || item.Status == "Clarification"? false : true;
         // button = <button disabled={isEditBtnDisable} type="button" className='btn btn-primary' id="add-row" onClick={() => this.OpenEditForm("EditMySubmission", item)}>Edit</button>
          button = <button disabled={isEditBtnDisable} type="button"  id="add-row" onClick={() => this.OpenEditForm("EditMySubmission", item)}><i className="fa fa-pencil-square-o" title='Edit'></i></button>
        }
        if (this.state.SelectedTabType == "MyTask") {
          //button = <button disabled={isEditBtnDisable} type="button" className='btn btn-primary' id="add-row" onClick={() => this.OpenEditForm("EditMySubmission", item)}>Action</button>
          button = <button disabled={isEditBtnDisable} type="button" id="add-row" onClick={() => this.OpenEditForm("EditMySubmission", item)}><i className="fa fa-tasks" title='Action' ></i></button>
        }
        return <span>
          {button}
        </span>;
      }
    },
  
    {
      name: "View",
      displayName: "",
      minWidth: 35,
      maxWidth: 35,
      render: (item: any) => {
        //return <div className='btn btn-primary' id="add-row" onClick={() => this.OpenEditForm("ViewMySubmission", item)}>View</div>;
        return   <button  type="button" id="add-row" onClick={() => this.OpenEditForm("ViewMySubmission", item)}><i className="fa fa-eye" title='View' ></i></button>
      }
    },
    {
      name: this.state.SelectedTabType == "MySubmission" ? "Progress" : "",
      displayName: "",
      minWidth: 60,
      maxWidth: 60,
      render: (item: any) => {
        let statusIcon;
        if ((item.Status=="Rejected" || item.Status=="Rejected by Accountant" || item.Status=="Rejected by HighManagement")) {
         // statusIcon = <i className="fa fa-book fa-f" onClick={() => this.OpenStatusBarPopUp(item)} title='Show Progress'></i> 
         
         if (this.state.SelectedTabType == "MySubmission") {
        // statusIcon=   <img title="Click here to see progress" onClick={() => this.OpenStatusBarPopUp(item)} className={contentEditFormStyles.statusImg} src={this.props.siteUrl + '/SiteAssets/ICONS/decline.png'} />
        statusIcon = <button  type="button"  id="add-row" onClick={() => this.OpenStatusBarPopUp(item)}><img title={item.Status} className={contentEditFormStyles.statusImg} src={this.props.siteUrl + '/SiteAssets/ICONS/decline.png'} /></button>
      }
        }
         if ((item.Status=="InProgress" || item.Status=="Approved" || item.Status=="Submitted" || item.Status=="Pending for Accountant")) {
          if (this.state.SelectedTabType == "MySubmission") {
          //statusIcon=   <img title="Click here to see progress" onClick={() => this.OpenStatusBarPopUp(item)} className={contentEditFormStyles.statusImg} src={this.props.siteUrl + '/SiteAssets/ICONS/timer.png'} />
          statusIcon = <button  type="button"  id="add-row" onClick={() => this.OpenStatusBarPopUp(item)}><img title={item.Status} className={contentEditFormStyles.statusImg} src={this.props.siteUrl + '/SiteAssets/ICONS/timer.png'} /></button>
        }
        }
         if (item.Status=="Order Placed") {
          if (this.state.SelectedTabType == "MySubmission") {
          //statusIcon=   <img title="Click here to see progress" onClick={() => this.OpenStatusBarPopUp(item)} className={contentEditFormStyles.statusImg} src={this.props.siteUrl + '/SiteAssets/ICONS/checked.png'} />
          statusIcon = <button  type="button"  id="add-row" onClick={() => this.OpenStatusBarPopUp(item)}><img title={item.Status} className={contentEditFormStyles.statusImg} src={this.props.siteUrl + '/SiteAssets/ICONS/checked.png'} /></button> 
        }
        }
         if (item.Status=="Clarification") {
          if (this.state.SelectedTabType == "MySubmission") {
          //statusIcon=   <img title="Click here to see progress" onClick={() => this.OpenStatusBarPopUp(item)}  className={contentEditFormStyles.statusImg} src={this.props.siteUrl + '/SiteAssets/ICONS/clear.png'} />
          statusIcon = <button  type="button"  id="add-row" onClick={() => this.OpenStatusBarPopUp(item)}><img title={item.Status} className={contentEditFormStyles.statusImg} src={this.props.siteUrl + '/SiteAssets/ICONS/clear.png'} /></button> 
        }
        }
         if (item.Status=="Draft") {
          if (this.state.SelectedTabType == "MySubmission") {
          statusIcon = <button  type="button"  id="add-row" onClick={() => this.OpenStatusBarPopUp(item)}><img title={item.Status} className={contentEditFormStyles.statusImg} src={this.props.siteUrl + '/SiteAssets/ICONS/notepad.png'} /></button> 
        }
        }
        return <span>
          {statusIcon}
        </span>;
      }
    },
    {
      name: "BlanketPORequest",
      displayName: "BlanketPORequest",
      isResizable: true,
      sorting: true,
      minWidth: 150,
      maxWidth: 150,
      render: (item: any) => {
        return <a href="javascript:void(0)" onClick={() => this.OpenViewFrom(item)}>{item.BlanketPORequest}</a>;
      }
    },
    {
      name: "RequestorID",
      displayName: "Request ID",
      isResizable: true,
      sorting: true,
      minWidth: 100,
      maxWidth: 100,
      render: (item: any) => {
        return <a href="javascript:void(0)" onClick={() => this.OpenViewFrom(item)}>{item.RequestorID}</a>;
      }
    },
    // {
    //   name: "Phone",
    //   displayName: "Phone",
    //   isResizable: true,
    //   sorting: true,
    //   minWidth: 90,
    //   maxWidth:90,
    //   render: (item: any) => {
    //     return <a href="javascript:void(0)" onClick={() => this.OpenViewFrom(item)}>{item.Phone}</a>;
    //   }
    // },

    // {
    //   name: "DateRequired",
    //   displayName: "Date Required",
    //   isResizable: true,
    //   sorting: true,
    //   minWidth: 90,
    //   maxWidth: 100,
    //   render: (item: any) => {
    //     return <a href="javascript:void(0)" onClick={() => this.OpenViewFrom(item)}>{item.DateRequired}</a>;
    //   }
    // },
    // {
    //   name: "ShipAddress",
    //   displayName: "Ship Address",
    //   isResizable: true,
    //   sorting: true,
    //   minWidth: 100,
    //   maxWidth: 140,
    //   render: (item: any) => {
    //     return <a href="javascript:void(0)" onClick={() => this.OpenViewFrom(item)}>{item.ShipAddress}</a>;
    //   }
    // },
 
    // {
    //   name: "PaymentFrom",
    //   displayName: "PaymentFrom",
    //   isResizable: true,
    //   sorting: true,
    //   minWidth: 90,
    //   maxWidth: 100,
    //   render: (item: any) => {
    //     return <a href="javascript:void(0)" onClick={() => this.OpenViewFrom(item)}>{item.PaymentFrom}</a>;
    //   }
    // },
    {
      name: "TotalExpense",
      displayName: "Amount($)",
      isResizable: true,
      sorting: true,
      minWidth: 100,
      maxWidth: 100,
      render: (item: any) => {
        return <a href="javascript:void(0)" onClick={() => this.OpenViewFrom(item)}>{item.TotalExpense}</a>;
      }
    },
    {
      name: "Status",
      displayName: "Status",
      isResizable: true,
      sorting: true,
      minWidth: 200,
      maxWidth: 200,
      render: (item: any) => {
        // let statusHtml;
        //   statusHtml = <span><a href="javascript:void(0)" onClick={() => this.OpenViewFrom(item)}>{item.Status}</a></span>;
        // return <span>
        //   {statusHtml}
        // </span>;

          let statusHtml;
          if (item.Status == "Order Placed") {
            statusHtml = <span><a href="javascript:void(0)" onClick={() => this.OpenViewFrom(item)}>{item.Status}</a><div>({item.OrderedDate})</div></span>
          } else {
            statusHtml = <span><a href="javascript:void(0)" onClick={() => this.OpenViewFrom(item)}>{item.Status}</a></span>;
          }
          return <span>
            {statusHtml}
          </span>;
      }
    },
    {
      name:  "Submitted By",
      displayName: "",
      minWidth: 100,
      maxWidth: 100,
      render: (item: any) => {      
        return <a href="javascript:void(0)" onClick={() => this.OpenViewFrom(item)}>{item.Creator}</a>;
        
      }
    },
    {
      name: "Approved By",
      displayName: "",
      isResizable: true,
      sorting: true,
      minWidth: 100,
      maxWidth: 100,
      render: (item: any) => {
        let statusHtml;
        if (item.ManagerStatus == "Approved" && item.ReviewForHighManagement =="Yes") {
          statusHtml = <span><a href="javascript:void(0)" onClick={() => this.OpenViewFrom(item)}>{item.Manager}</a></span>
        } 
        return <span>
          {statusHtml}
        </span>;
      }
    },
    // {
    //   name: this.state.SelectedTabType == "MySubmission" ? "" : this.state.SelectedTabType == "MyTask" ? "Action" : "",
    //   displayName: "",
    //   minWidth: 60,
    //   maxWidth: 60,
    //   render: (item: any) => {
    //     let isEditBtnDisable = false;
    //     let button;
    //     if (this.state.SelectedTabType == "MySubmission") {
    //       isEditBtnDisable = item.Status == "Rejected by HighManagement" || item.Status == "Rejected by Accountant" || item.Status == "Rejected" || item.Status == "Draft" || item.Status == "Clarification"? false : true;
    //       button = <button disabled={isEditBtnDisable} type="button" className='btn btn-primary' id="add-row" onClick={() => this.OpenEditForm("EditMySubmission", item)}>Edit</button>
    //     }
    //     if (this.state.SelectedTabType == "MyTask") {
    //       button = <button disabled={isEditBtnDisable} type="button" className='btn btn-primary' id="add-row" onClick={() => this.OpenEditForm("EditMySubmission", item)}>Action</button>
    //     }
    //     return <span>
    //       {button}
    //     </span>;
    //   }
    // },
  
    // {
    //   name: "",
    //   displayName: "",
    //   minWidth: 60,
    //   maxWidth: 60,
    //   render: (item: any) => {
    //     return <div className='btn btn-primary' id="add-row" onClick={() => this.OpenEditForm("ViewMySubmission", item)}>View</div>;
    //   }
    // },
    ];
    return viewFields;
  };
  public groupByFields() {
    const groupByFields: IGrouping[] = [
      {
        name: "Status",
        order: GroupOrder.descending
      }
    ];
    return groupByFields
  };

  //FOR Update
    //* Add Procurement details in ProcurementDetails list*/
    addUpdateAllProcurementDetails = (itemId, submissionType: string) => {
      if (this.state.IProcurementModel.length > 0) {
        let procurementItemsCreate = [];
        let procurementItemsUpdate = [];
        this.state.IProcurementModel.map((procurement) => {
          if (procurement.Id != undefined) {
            procurementItemsUpdate.push(procurement);
          } else {
            procurementItemsCreate.push(procurement);
          }
        })
        if (procurementItemsCreate.length > 0 && procurementItemsUpdate.length > 0) {
          this._addProcurementDetailsAsBatch(procurementItemsCreate, itemId).then(() => {
            this.UpdateProcurementDetailsAsBatch(procurementItemsUpdate).then(() => {
              alert("Request " + submissionType + " sucessfully");
              $('#loaderEdit').hide();
              this.RefreshPage();
            })
  
          });
        }
        if (procurementItemsCreate.length > 0 && procurementItemsUpdate.length == 0) {
          this._addProcurementDetailsAsBatch(procurementItemsCreate, itemId).then(() => {
            alert("Request " + submissionType + " sucessfully");
            $('#loaderEdit').hide();
            this.RefreshPage();
          })
        }
        if (procurementItemsUpdate.length > 0 && procurementItemsCreate.length == 0) {
          this.UpdateProcurementDetailsAsBatch(procurementItemsUpdate).then(() => {
            alert("Request " + submissionType + " sucessfully");
            $('#loaderEdit').hide();
            this.RefreshPage();
            //alert(submissionType=="Submited"?"Request submitted sucessfully":"Request drafted sucessfully");
          })
        }
      }
      else {
        alert("Request " + submissionType + " sucessfully");
        $('#loaderEdit').hide();
        this.RefreshPage();
      }
  
    }
    private isModalOpen(): boolean {
      return this.state.openDialog;
    };
    private isStatusBarModalOpen(): boolean {
      return this.state.openStatusBarDialog;
    };
    private hideStatusBarModal=()=>{
      this.setState({
        openStatusBarDialog:false
      })
      this.RefreshPage();
    };
    RefreshPage = () => {
      //window.location.reload();
      this.setState({
        IProcurementModel: [],
        Comments: "",
        Phone: "",
        BlanketPORequest: "",
        PaymentFrom: "",
        DateRequired: null,
        ShipAddress: "",
        PreferredVendorOptions:[],
        PaymentFromOptions:[],
        ShipAddressOptions:[],
        IsBtnClicked: false,
        //SelectedTabType:"",
        fileInfos: [],
        ManagerId: null,
        filePickerResult: [],
        Attachments: [],
        SelectedFiles: [],
        CurrentUserName:"",
        openDialog: false,
        openEditDialog: false,
        HighManagementIds:null,
        selectedProcurement: {},
        latestComments:"",
        Manager:"",
        Creator:"",
        IsHighManagement:false,
        IsManager:false,
        ManagerStatus:"",
        PurchasingEntity:"",
        PayingEntity:"",
        PurchasingEntityOptions:[],
        PayingEntityOptions:[],
        MumbaiAddressDetails:"",
        IndoreAddressDetails:"",
        DelhiAddressDetails:"",
        NoidaAddressDetails:"",
        OtherAddressDetails:"",
        HighManagementStatus:"",
        AccountantStatus:"",
        openStatusBarDialog:false,
        HighManagement:"",
        Accountant:"",
        RequestorResponse:"",
        ManagerResponse:"",
        HighManagementResponse:"",
        AccountantResponse:"",
      })
      this.componentDidMount();
    }
    //Update Procurement
    UpdateProcurementRequest = (submissionType) => {
      var isError = this.ValidateForm(submissionType);
      if (!isError) {
        this.setState({
          IsBtnClicked: true,
        })
        $('#loaderEdit').show();
        let previousComments = this.state.Comments == null ? "" : this.state.Comments;
        let todayDate = this._spOps.GetTodaysDate();
        let latestCommentsHTML = '<strong>' + this.state.CurrentUserName + ' : ' + todayDate + '</strong>' + '<div>' + this.state.latestComments + '</div>';
        // let totalAmount: any = 0;
        // if (this.state.IProcurementModel.length > 0) {
        //   this.state.IProcurementModel.map((expense) => {
        //     totalAmount += expense.TotalCost != undefined ? parseFloat(expense.TotalCost) : 0;
        //   })
        // }
        let totalAmount: any = 0;
         totalAmount=this._calculateTotalCost();
        let updatePostData: any = {};
        updatePostData = {
          Title: this.state.BlanketPORequest,
          Phone: this.state.Phone,
          DateRequired: this.state.DateRequired != "" ? this.state.DateRequired : null,
          PaymentFrom: this.state.PaymentFrom,
          ShipAddress:this.state.ShipAddress,
          OtherAddress:this.state.OtherAddress,
          PurchasingEntity:this.state.PurchasingEntity,
          PayingEntity:this.state.PayingEntity,
          //FinanceId: this.state.FinanceId,
          TotalExpense: totalAmount.toString(),
          
        }
        if (this.state.latestComments != "") {
          updatePostData.Comments = latestCommentsHTML.concat(previousComments);
        }
        if (this.state.SelectedTabType == "MySubmission") {
          if (submissionType == "Submitted") {
            // updatePostData.ManagerStatus = "Pending for Manager";
            // updatePostData.Status = "Submitted";
            // updatePostData.ManagerId = this.state.ManagerId;

           //if manager is not null
           if(this.state.ManagerId!=null){
          updatePostData.ManagerStatus = this.state.ManagerId!=null?"Pending for Manager":"Manager Approval Not Required";
          updatePostData.Status = this.state.ManagerId!=null?"Submitted":"InProgress";
          updatePostData.HighManagementStatus = this.state.ManagerId!=null?"":"Pending for HighManagement";
          if(this.state.AccountantStatus=="Clarification" && this.state.ManagerId!=null){
            updatePostData.HighManagementStatus = "Approved";
          }
          if(this.state.AccountantStatus=="Clarification" && this.state.HighManagementStatus=="HighManagement Approval Not Required" && this.state.ManagerId!=null){
            updatePostData.HighManagementStatus = "HighManagement Approval Not Required";
          }
          if(this.state.AccountantStatus=="Rejected" && this.state.ManagerId!=null){
            updatePostData.HighManagementStatus = this.state.HighManagementStatus;
          }
          if(this.state.HighManagementStatus=="Rejected" && this.state.ManagerId!=null){
            updatePostData.HighManagementStatus = this.state.HighManagementStatus;
          }
          updatePostData.ReviewForHighManagement = this.state.ManagerId!=null?"":"Yes";
          updatePostData.ManagerId = this.state.ManagerId;
           }
          //if manager is null and amount > 2500
           if(totalAmount > parseInt(this.state.ManagerApprovalCost) && this.state.ManagerId==null){
            updatePostData.ManagerStatus = "Manager Approval Not Required";
            updatePostData.Status = "InProgress";
            updatePostData.HighManagementStatus = "Pending for HighManagement";
            updatePostData.ReviewForHighManagement = "Yes";
            updatePostData.HighManagementId= {"results":this.state.HighManagementIds};
           }
           //if manager is null and amount < 2500
           if(totalAmount < parseInt(this.state.ManagerApprovalCost) && this.state.ManagerId==null){
            updatePostData.ManagerStatus = "Manager Approval Not Required";
            updatePostData.Status = "Pending for Accountant";
            updatePostData.HighManagementStatus = "HighManagement Approval Not Required";
            updatePostData.ReviewForHighManagement = "No";
            updatePostData.AccountantStatus="Pending for Accountant";
            updatePostData.AccountantId = { "results": this.state.AllAccoutantIds };
           }

           if (this.state.selectedProcurement.AccountantStatus == "Clarification" && this.state.ManagerId != null) {
            updatePostData.ManagerStatus = "Approved";
            updatePostData.AccountantStatus = "Pending for Accountant";
            updatePostData.Status = "Pending for Accountant";
            //updatePostData.ReviewForFinace = "Yes";
          }
          if ((this.state.selectedProcurement.AccountantStatus == "Clarification" && this.state.ManagerId == null) || (this.state.selectedProcurement.AccountantStatus == "Rejected" && this.state.ManagerId == null)) {
            updatePostData.ManagerStatus = "Manager Approval Not Required";
            updatePostData.AccountantStatus = "Pending for Accountant";
            updatePostData.Status = "Pending for Accountant";
            //updatePostData.ReviewForFinace = "Yes";
          }
          if((this.state.selectedProcurement.AccountantStatus == "Clarification" || this.state.selectedProcurement.AccountantStatus == "Rejected") && totalAmount > parseInt(this.state.ManagerApprovalCost) && this.state.ManagerId==null){
            updatePostData.ManagerStatus = "Manager Approval Not Required";
            updatePostData.Status = "InProgress";
            updatePostData.HighManagementStatus = "Pending for HighManagement";
            updatePostData.ReviewForHighManagement = "Yes";
            updatePostData.HighManagementId= {"results":this.state.HighManagementIds};
           }
           //if manager is null and amount < 2500
           if((this.state.selectedProcurement.AccountantStatus == "Clarification" || this.state.selectedProcurement.AccountantStatus == "Rejected") && totalAmount < parseInt(this.state.ManagerApprovalCost) && this.state.ManagerId==null){
            updatePostData.ManagerStatus = "Manager Approval Not Required";
            updatePostData.Status = "Pending for Accountant";
            updatePostData.HighManagementStatus = "HighManagement Approval Not Required";
            updatePostData.ReviewForHighManagement = "No";
            updatePostData.AccountantStatus="Pending for Accountant";
            updatePostData.AccountantId = { "results": this.state.AllAccoutantIds };
           }
            // if (this.state.ManagerId!=null) {
            //   updatePostData.ManagerStatus = "Approved";
            //   updatePostData.HighManagementStatus = "Pending for HighManagement";
            //   updatePostData.Status = "InProgress";
            // }
            // if (this.state.selectedProcurement.HighManagementStatus == "Rejected" && this.state.ManagerId==null) {
            //   updatePostData.ManagerStatus = "Manager Approval Not Required";
            //   updatePostData.HighManagementStatus = "Pending for HighManagement";
            //   updatePostData.Status = "InProgress";
            // }
            
            
          }
          if (submissionType == "Draft") {
            updatePostData.ManagerStatus = "";
            updatePostData.Status = "Draft";
            updatePostData.ManagerId = this.state.ManagerId;
          }
        }
        if (this.state.SelectedTabType == "MyTask" && this.state.IsManager) {
         
          if (submissionType == "Rejected") {
            updatePostData.Status = "Rejected";
            updatePostData.ManagerStatus = submissionType;
          }
          if (submissionType == "Rejected" && (this.state.ManagerStatus=="Approved" || this.state.ManagerStatus=="Manager Approval Not Required")) {
            updatePostData.Status = "Rejected";
            updatePostData.ManagerStatus = this.state.ManagerStatus;
          }

          if (submissionType == "Approved") {
            if(totalAmount>parseInt(this.state.ManagerApprovalCost)){
            updatePostData.ReviewForHighManagement = "Yes";
            updatePostData.Status = "InProgress";
            updatePostData.HighManagementStatus = "Pending for HighManagement";
            updatePostData.ManagerStatus = submissionType;
            updatePostData.HighManagementId= {"results":this.state.HighManagementIds};
            }
            if(totalAmount < parseInt(this.state.ManagerApprovalCost)){
            updatePostData.HighManagementStatus = "HighManagement Approval Not Required";
            updatePostData.Status = "Pending for Accountant";
            updatePostData.ReviewForHighManagement = "No";
            updatePostData.AccountantStatus = "Pending for Accountant";
            updatePostData.AccountantId = { "results": this.state.AllAccoutantIds };
            updatePostData.ManagerStatus = submissionType;
            }
          }
        }
        if (this.state.SelectedTabType == "MyTask" && this.state.IsHighManagement) {
          
          if (submissionType == "Approved") {
            updatePostData.Status = "Pending for Accountant";
            updatePostData.AccountantStatus = "Pending for Accountant";
            updatePostData.AccountantId = { "results": this.state.AllAccoutantIds };
            updatePostData.HighManagementStatus = submissionType;
          }
          
          
          if (submissionType == "Rejected") {
           // updatePostData.ManagerStatus = "Rejected by HighManagement";
            updatePostData.Status = "Rejected by HighManagement";
            updatePostData.HighManagementStatus = submissionType;
          }
        }
        if (this.state.SelectedTabType == "MyTask" && this.state.IsAccountant) {
          //updatePostData.AccountantStatus = submissionType;
         // updatePostData.Status = submissionType;
           if (submissionType == "Order Placed") {
            updatePostData.AccountantStatus = "Order Placed";
             updatePostData.Status = "Order Placed";
             updatePostData.OrderedDate = this.state.OrderedDate;
           }
           if (submissionType == "Rejected" && (this.state.ManagerStatus=="Approved" || this.state.ManagerStatus=="Manager Approval Not Required")) {
           // updatePostData.ManagerStatus = "Rejected by Accountant";
            //updatePostData.HighManagementStatus = this.state.HighManagement==null || this.state.HighManagement=="" ?"": "Rejected by Accountant";
            updatePostData.Status = "Rejected by Accountant";
            updatePostData.AccountantStatus ="Rejected";
          }
          if(submissionType=="Clarification"){
            updatePostData.AccountantStatus = "Clarification";
            updatePostData.Status = "Clarification";
          }
        }
        let logHistoryPostData:any={};
      logHistoryPostData={
        Title:this.state.selectedProcurement.RequestorID,
        ProcurementId:this.state.selectedProcurement.ID,
        CommentsHistory:this.state.latestComments,
        Status:updatePostData.Status,
       // NameId:this.state.CurrentUserID
      }
      this._spOps.CreateItem(this.props.logHistoryListTitle, logHistoryPostData).then((result: any) => {});
        this._spOps.UpdateItem(this.props.procurementRequestList, updatePostData, this.state.selectedProcurement.ID).then((result: any) => {
          console.log(this.state.selectedProcurement.ID);
          let itemId = this.state.selectedProcurement.ID;
          let { fileInfos } = this.state;
          let web = Web(this.props.siteUrl);
          if (fileInfos.length > 0) {
            let fileToAttach = [];
            fileInfos.map((fileItem) => {
              if (fileItem.ServerRelativeUrl == undefined) {
                fileToAttach.push(fileItem);
              }
            })
  
            web.lists.getByTitle(this.props.procurementRequestList).items.getById(itemId).attachmentFiles.addMultiple(fileToAttach).then(() => {
              this.addUpdateAllProcurementDetails(itemId, submissionType);
            });
          }
          else {
            this.addUpdateAllProcurementDetails(itemId, submissionType);
          }
          if (this.state.FilesToDelete.length) {
            web.lists.getByTitle(this.props.procurementRequestList).items.getById(this.state.selectedProcurement.ID).attachmentFiles.deleteMultiple(...this.state.FilesToDelete);
          }
          if (this.state.ProcurementItemsToDelete.length > 0) {
            this.DeleteProcurementDetails(this.state.ProcurementItemsToDelete);
          }
        })
  
      }
    }
    // async function with await to delete procurement from procurementReqDetailList
    private async DeleteProcurementDetails(Procurements: any[]) {
      let web = Web(this.props.siteUrl);
      for (const procurement of Procurements) {
        await web.lists.getByTitle(this.props.procurementReqDetailList).items.getById(procurement.Id).delete();
      }
    }
    //** call validation method and create item into list */
    UpdateRequest = (submissionType) => {
      if (submissionType == "Rejected" && confirm('Are you sure, you want to reject')) {
        this.UpdateProcurementRequest(submissionType);
      } else {
        this.UpdateProcurementRequest(submissionType);
      }
    }
  
    /* This event will fire on remove specific row */
    handleRemoveSpecificRow = (idx) => () => {
      if (confirm('Are you sure, you want to remove it ?')) {
        try {
          const rows = this.state.IProcurementModel
          let procurementItems = [];
          this.state.IProcurementModel.map((item, index) => {
            if (item.Id != undefined && index === idx) {
              procurementItems.push(item);
            }
          })
          if (rows.length > 1) {
            rows.splice(idx, 1);
          }
          let tempItemsToDeleteArr = this.state.ProcurementItemsToDelete;
          tempItemsToDeleteArr.push.apply(tempItemsToDeleteArr, procurementItems);
          this.setState({ IProcurementModel: rows, ProcurementItemsToDelete: tempItemsToDeleteArr });
        } catch (error) {
          console.log("Error in React Table handle Remove Specific Row : " + error);
        }
      }
    };
    //** onClick of View Expense, Generate html for PDF*/
  renderInvoiceTableData() {
    return this.state.IProcurementModel.map((item, idx) => {
      return (<tbody key={idx}>
                    <tr>
                      <td className={contentInvoiceStyles.td}>{this.state.IProcurementModel[idx].PreferredVendor}</td>
                      <td className={contentInvoiceStyles.td}>{this.state.IProcurementModel[idx].ItemDescription} </td>
                      <td className={contentInvoiceStyles.td}>{this.state.IProcurementModel[idx].Quantity} </td>
                      <td className={contentInvoiceStyles.td}>{this.state.IProcurementModel[idx].UnitCost} </td>
                      {/* <td className={contentInvoiceStyles.td}>{this.state.IProcurementModel[idx].StartMile} </td>
                      <td className={contentInvoiceStyles.td}>{this.state.IProcurementModel[idx].EndMileout} </td>
                      <td className={contentInvoiceStyles.td}>{this.state.IProcurementModel[idx].Description}</td> */}
                      <td className={contentInvoiceStyles.td}>{this.state.IProcurementModel[idx].TotalCost}</td>
                    </tr>
                  </tbody>)               
    })
  }
//** onClick of View Expense, Generate html for PDF*/
renderInvoiceCommentTableData() {
  return this.state.ILogHistoryModel.map((item, idx) => {
    return (<tbody key={idx}>
                  <tr>
                    <td className={contentInvoiceStyles.td}>{this.state.ILogHistoryModel[idx].CommentsHistory}</td>
                    <td className={contentInvoiceStyles.td}>{this.state.ILogHistoryModel[idx].Author} </td>
                    <td className={contentInvoiceStyles.td}>{this.state.ILogHistoryModel[idx].Status} </td>
                    <td className={contentInvoiceStyles.td}>{this.state.ILogHistoryModel[idx].CreatedOn} </td>
                  </tr>
                </tbody>)               
  });
};
    // Render Procurement Table
    renderTableData() {
      var selectHeight = {
        color: 'black',
        'margin-top': '-4px',
      };
      return this.state.IProcurementModel.map((item, idx) => {
        return (<div key={idx}>
          <div className={styles.renderProcurementTbl}>
            <div className="form-group col-md-2">
            {idx==0 && <span>  <label className={styles.lblCtrl}>Preferred Vendor</label><span className={styles.star}>*</span></span>}
              <select className='form-control' style={selectHeight} name="PreferredVendor" value={this.state.IProcurementModel[idx].PreferredVendor} id={this.state.IProcurementModel[idx].id} onChange={this._handleChange(idx)}>
                <option value="">Select</option>
                {this._renderDropdown(this.state.PreferredVendorOptions)}
              </select>
              {item.isPreferredVendorError == true && <span className={styles.errMsg}>{item.PreferredVendorErrMsg}</span>}
            </div>
              <div className="form-group col-md-2">
              {idx==0 && <span> <label className="control-label">Item Description</label><span className={styles.star}>*</span></span>}
                <input
                  placeholder='Item Description'
                  type="text"
                  className='form-control'
                  name="ItemDescription"
                  value={this.state.IProcurementModel[idx].ItemDescription}
                  onChange={this._handleChange(idx)}
                  id={this.state.IProcurementModel[idx].id}
                />
  
  {item.isItemDescriptionError == true && <span className={styles.errMsg}>{item.ItemDescriptionErrMsg}</span>}
              </div>
              <div className="form-group col-md-2">
              {idx==0 && <span> <label className="control-label">Link to the site</label><span className={styles.star}>*</span></span>}
                <input
                  placeholder='Link to the site'
                  type="text"
                  className='form-control'
                  name="SiteLink"
                  value={this.state.IProcurementModel[idx].SiteLink}
                  onChange={this._handleChange(idx)}
                  id={this.state.IProcurementModel[idx].id}
                />
               {item.isSiteLinkError == true && <span className={styles.errMsg}>{item.SiteLinkErrMsg}</span>} 
              </div>

              <div className="form-group col-md-1">
               {idx==0 &&  <label className="control-label">Quantity <span className={styles.star}>*</span></label>}
                <input
                  placeholder='Quantity'
                  type="number"
                  min="1"
                  className='form-control'
                  name="Quantity"
                  value={this.state.IProcurementModel[idx].Quantity}
                  onChange={this._handleChange(idx)}
                  id={this.state.IProcurementModel[idx].id}
                />
                  {item.isQuantityError == true && <span className={styles.errMsg}>{item.QuantityErrMsg}</span>}
              </div>
              <div className="form-group col-md-2">
                 {idx==0 && <label className="control-label">$Unit Cost <span className={styles.star}>*</span></label>}
                <input
                  placeholder='Unit Cost'
                  type="number"
                  min="1"
                  className='form-control'
                  name="UnitCost"
                  value={this.state.IProcurementModel[idx].UnitCost}
                  onChange={this._handleChange(idx)}
                  id={this.state.IProcurementModel[idx].id}
                />
                  {item.isUnitCostError == true && <span className={styles.errMsg}>{item.UnitCostErrMsg}</span>}
              </div>

              <div className="form-group col-md-2">
                {idx==0 && <label className="control-label">$Total Cost</label>}
                <input
                  placeholder='Total Cost'
                  type="text"
                  className='form-control'
                  disabled={true}
                  name="TotalCost"
                  value={this.state.IProcurementModel[idx].TotalCost}
                  onChange={this._handleChange(idx)}
                  id={this.state.IProcurementModel[idx].id}
                />
              </div>

            {this.state.IProcurementModel.length > 1 &&
              <div className="form-group col-md-1">
               {idx==0 &&  <label className="control-label"></label>}
                <div onClick={this._handleRemoveSpecificRow(idx)} className={styles.deleteIcon}>
                  <Icon iconName="delete" className="ms-IconExample" />
                </div>
              </div>
            }
  
          </div>
  
        </div>)
      })
    }
    //** onClick of Add Expense, Render Expense details fields in table*/
  renderTableDataEdit() {
    let isProcurementDeleteIcon = this.state.SelectedTabType != "MyTask" && this.state.selectedProcurement.formType == "Edit" ? true : false;
    return this.state.IProcurementModel.map((item, idx) => {
      return (<span key={idx}>
        {/* <div className={styles.expenselbl}>Expense {idx+1}</div> */}
        <div className={this.state.SelectedTabType != "MyTask" && this.state.selectedProcurement.formType == "Edit" ? contentEditFormStyles.formRow : contentEditFormStyles.editRequestTable}>
        <div className="form-group col-md-2">
            {idx==0 && <span>  
              <label className={contentEditFormStyles.lblCtrl}>Preferred Vendor<span className={contentEditFormStyles.star}>*</span></label></span>}
              <select className='form-control' disabled={this.state.SelectedTabType == "MyTask" || this.state.selectedProcurement.formType == "View" ? true : false} name="PreferredVendor" value={this.state.IProcurementModel[idx].PreferredVendor} id={this.state.IProcurementModel[idx].id} onChange={this._handleChange(idx)}>
                <option value="">Select</option>
                {this._renderDropdown(this.state.PreferredVendorOptions)}
              </select>
              {item.isPreferredVendorError == true && <span className={contentEditFormStyles.errMsg}>{item.PreferredVendorErrMsg}</span>}
            </div>
              <div className="form-group col-md-2"> {idx==0 && <span>
              <label className={contentEditFormStyles.lblCtrl}>Item Description<span className={contentEditFormStyles.star}>*</span></label></span>}
                <input
                  placeholder='Item Description'
                  type="text"
                  className='form-control'
                  disabled={this.state.SelectedTabType == "MyTask" || this.state.selectedProcurement.formType == "View" ? true : false}
                  name="ItemDescription"
                  value={this.state.IProcurementModel[idx].ItemDescription}
                  onChange={this._handleChange(idx)}
                  id={this.state.IProcurementModel[idx].id}
                />
  
                {item.isItemDescriptionError == true && <span className={contentEditFormStyles.errMsg}>{item.ItemDescriptionErrMsg}</span>}
              </div>
              <div className="form-group col-md-2">
              {idx==0 && <span> 
                <label className={contentEditFormStyles.lblCtrl}>Link to the Site<span className={contentEditFormStyles.star}>*</span></label>
                
                </span>}
                <div className="">
                <input
                  placeholder='Link to the site'
                  type="text"
                  className='form-control'
                  disabled={this.state.SelectedTabType == "MyTask" || this.state.selectedProcurement.formType == "View" ? true : false}
                  name="SiteLink"
                  value={this.state.IProcurementModel[idx].SiteLink}
                  onChange={this._handleChange(idx)}
                  id={this.state.IProcurementModel[idx].id}
                />
                {this.state.selectedProcurement.formType == "View" && <span className={contentEditFormStyles.link}><a href={this.state.IProcurementModel[idx].SiteLink} target="_blank" data-interception="off"> 
                  <img title="Click here to open link" className={contentEditFormStyles.linkImg} src={this.props.siteUrl + '/SiteAssets/ICONS/Link.png'} /></a>
                  </span>}
                </div>
                
             
                {item.isSiteLinkError == true && <span className={contentEditFormStyles.errMsg}>{item.SiteLinkErrMsg}</span>}
              </div>

              <div className="form-group col-md-1">
               {idx==0 &&  <label className={contentEditFormStyles.lblCtrl}>Quantity<span className={contentEditFormStyles.star}>*</span></label>}
                <input
                  placeholder='Quantity'
                  type="number"
                  min="1"
                  className='form-control'
                  disabled={this.state.SelectedTabType == "MyTask" || this.state.selectedProcurement.formType == "View" ? true : false}
                  name="Quantity"
                  value={this.state.IProcurementModel[idx].Quantity}
                  onChange={this._handleChange(idx)}
                  id={this.state.IProcurementModel[idx].id}
                />
                 {item.isQuantityError == true && <span className={contentEditFormStyles.errMsg}>{item.QuantityErrMsg}</span>}
              </div>
              <div className="form-group col-md-2">
                 {idx==0 && <label className={contentEditFormStyles.lblCtrl}>$Unit Cost<span className={contentEditFormStyles.star}>*</span></label>}
                <input
                  placeholder='Unit Cost'
                  type="number"
                  min="1"
                  className='form-control'
                  disabled={this.state.SelectedTabType == "MyTask" || this.state.selectedProcurement.formType == "View" ? true : false}
                  name="UnitCost"
                  value={this.state.IProcurementModel[idx].UnitCost}
                  onChange={this._handleChange(idx)}
                  id={this.state.IProcurementModel[idx].id}
                />
                {item.isUnitCostError == true && <span className={contentEditFormStyles.errMsg}>{item.UnitCostErrMsg}</span>}
              </div>

              <div className="form-group col-md-2">
                {idx==0 && <label className={contentEditFormStyles.lblCtrl}>$Total Cost</label>}
                <input
                  placeholder='Total Cost'
                  type="text"
                  className='form-control'
                  disabled={true}
                  name="TotalCost"
                  value={this.state.IProcurementModel[idx].TotalCost}
                  onChange={this._handleChange(idx)}
                  id={this.state.IProcurementModel[idx].id}
                />
              </div>
           {this.state.IProcurementModel.length > 1 &&
            <div className="form-group col-md-1">
             {idx==0 && <label></label>}
              <div id="delete-spec-row" className={contentEditFormStyles.deleteIcon}>
                {isProcurementDeleteIcon &&
                  <Icon iconName="delete" className="ms-IconExample" onClick={this.handleRemoveSpecificRow(idx)} />
                }
              </div>
            </div>
          } 
          <div className={contentEditFormStyles.itemLine}></div>
        </div>
        
      </span>)
    })
  }
  managerSectionCss(){
    let mngrCss="";
    switch (this.state.ManagerStatus) {
      case 'Approved':
        mngrCss = [contentStatusBarStyles.statuscircle, contentStatusBarStyles.approved].join(" ");
        break;
      case 'Manager Approval Not Required':
        mngrCss = [contentStatusBarStyles.statuscircle, contentStatusBarStyles.approved].join(" ");
        break;
      case 'Rejected':
        mngrCss = [contentStatusBarStyles.statuscircle, contentStatusBarStyles.rejected].join(" ");
        break;
      case 'Rejected by HighManagement':
        mngrCss = [contentStatusBarStyles.statuscircle, contentStatusBarStyles.rejected].join(" ");
        break;
      case 'Rejected by Accountant':
        mngrCss = [contentStatusBarStyles.statuscircle, contentStatusBarStyles.rejected].join(" ");
        break;
      case 'Pending for Manager':
        mngrCss = [contentStatusBarStyles.statuscircle,contentStatusBarStyles.pending].join(" ");
        break;
        default:
       mngrCss = [contentStatusBarStyles.statuscircle,""].join(" "); 
    }
    return mngrCss;
  };
  highmngmtSectionCss(){
    let coeCss="";
    switch (this.state.HighManagementStatus) {
      case 'Approved':
        coeCss = [contentStatusBarStyles.statuscircle, contentStatusBarStyles.approved].join(" ");
        break;
      case 'Clarification':
        coeCss = [contentStatusBarStyles.statuscircle, contentStatusBarStyles.clarify].join(" ");
        break;
      case 'Rejected':
        coeCss = [contentStatusBarStyles.statuscircle, contentStatusBarStyles.rejected].join(" ");
        break;
      case 'Pending':
        coeCss = [contentStatusBarStyles.statuscircle,contentStatusBarStyles.pending].join(" ");
        break;
      case 'Pending for Accountant':
        coeCss = [contentStatusBarStyles.statuscircle,contentStatusBarStyles.pending].join(" ");
        break;    
      case 'Pending for HighManagement':
        coeCss = [contentStatusBarStyles.statuscircle,contentStatusBarStyles.pending].join(" ");
        break;
      case 'HighManagement Approval Not Required':
        coeCss = [contentStatusBarStyles.statuscircle,contentStatusBarStyles.approved].join(" ");
        break;
      case 'Rejected by Accountant':
        coeCss = [contentStatusBarStyles.statuscircle, contentStatusBarStyles.rejected].join(" ");
        break;
      default:
        coeCss = [contentStatusBarStyles.statuscircle,""].join(" ");
    }
    return coeCss;
  }
  accountantSectionCss(){
    let accCss="";
    switch (this.state.AccountantStatus) {
      case 'Order Placed':
        accCss = [contentStatusBarStyles.statuscircle, contentStatusBarStyles.approved].join(" ");
        break;
      case 'Clarification':
        accCss = [contentStatusBarStyles.statuscircle, contentStatusBarStyles.clarify].join(" ");
        break;
      case 'Rejected':
        accCss = [contentStatusBarStyles.statuscircle, contentStatusBarStyles.rejected].join(" ");
        break;
      case 'Pending for Accountant':
        accCss = [contentStatusBarStyles.statuscircle,contentStatusBarStyles.pending].join(" ");
        break;
      default:
        accCss = [contentStatusBarStyles.statuscircle,""].join(" ");
    }
    return accCss;
  }
  public render(): React.ReactElement<IProcurementProps> {
    let managerCss="";
    let highManagementCss="";
    let accountantCss="";
    managerCss= this.managerSectionCss();
    highManagementCss=this.highmngmtSectionCss();
    accountantCss=this.accountantSectionCss();
    let isShowAddProcurementBtn = this.state.SelectedTabType != "MyTask" && this.state.selectedProcurement.formType == "Edit" ? true : false;
    let renderPendingTabs: any;
    let renderApprovedTabs: any;
    let renderRejectedTabs: any;
    let renderOrderPlacedTabs: any;
    if (this.state.IsHighManagement || this.state.IsManager || this.state.IsAccountant) {
      renderPendingTabs = <PivotItem itemKey="MyTask" key="MyTask" headerText="Pending" itemCount={this.state.MyPendingItems.length} onClick={() => this.selectedTabType("MyTask")}>
        <ListView
          items={this.state.MyPendingItems}
          showFilter={true}
          filterPlaceHolder="Search..."
          compact={true}
          //selectionMode={SelectionMode.single}
          onRenderRow={this.OnListViewRenderRow}
          listClassName={styles.listViewStyle}
          // selection={this.OpenViewFrom}    
          groupByFields={this.groupByFields()}
          viewFields={this.viewFields()}
        />
      </PivotItem>
    }
    if (this.state.IsHighManagement || this.state.IsManager) {
      renderApprovedTabs = <PivotItem itemKey="MyApprovedTask" key="MyApprovedTask" headerText="Approved" itemCount={this.state.MyApprovedItems.length} onClick={() => this.selectedTabType("MyApprovedTask")}>
        <ListView
          items={this.state.MyApprovedItems}
          showFilter={true}
          filterPlaceHolder="Search..."
          compact={true}
          //selectionMode={SelectionMode.single}
          onRenderRow={this.OnListViewRenderRow}
          listClassName={styles.listViewStyle}
          // selection={this.OpenViewFrom}  
          groupByFields={this.groupByFields()}
          viewFields={this.viewFields()}
        />
      </PivotItem>
    }
    if (this.state.IsAccountant) {
      renderOrderPlacedTabs = <PivotItem itemKey="MyPaidTask" key="MyPaidTask" headerText="Order Placed" itemCount={this.state.MyOrderedPlacedItems.length} onClick={() => this.selectedTabType("MyPaidTask")}>
        <ListView
          items={this.state.MyOrderedPlacedItems}
          showFilter={true}
          filterPlaceHolder="Search..."
          compact={true}
         // selectionMode={SelectionMode.single}
          onRenderRow={this.OnListViewRenderRow}
          listClassName={styles.listViewStyle}
          // selection={this.OpenViewFrom}    

          groupByFields={this.groupByFields()}
          viewFields={this.viewFields()}
        />
      </PivotItem>
    }

    if (this.state.IsHighManagement || this.state.IsManager) {
      renderRejectedTabs = <PivotItem itemKey="MyRejectedTask" key="MyRejectedTask" headerText="Rejected" itemCount={this.state.MyRejectedItems.length} onClick={() => this.selectedTabType("MyRejectedTask")}>
        <ListView
          items={this.state.MyRejectedItems}
          showFilter={true}
          filterPlaceHolder="Search..."
          compact={true}
          //selectionMode={SelectionMode.single}
          onRenderRow={this.OnListViewRenderRow}
          listClassName={styles.listViewStyle}
          // selection={this.OpenViewFrom}    
          groupByFields={this.groupByFields()}
          viewFields={this.viewFields()}
        />
      </PivotItem>
    }
    return (
      <div className={ styles.procurement }>
         <div className="row">
          <div className="form-group col-md-12">
            <h4 className={styles.headerLbl}>Procurement Management</h4>
          </div>
        </div>
           {/* <div className="">
          <div className={styles.headerContainer}><span className={styles.txtCenter}><h4>Procurement Request</h4></span>
           <div><strong>Inovatree pvt ltd</strong></div>
           <div><strong>Delhi</strong></div>
           <div><strong>110092</strong></div>
           </div>
          </div> */}
          <Pivot aria-label="OnChange Pivot Example" onLinkClick={this._onPivotItemClick}>
          <PivotItem itemKey="NewRequest" key="NewRequest" headerText="Submit new request">
          <div className={styles.groove}>
        <div className={styles.rowTable}>
            <div className="form-group col-md-3">
              <label className={styles.lblCtrl}>Phone</label><span className={styles.star}>*</span>
              <input
                placeholder='Phone'
                type="text"
                className='form-control'
                name="Phone"
                value={this.state.Phone}
                onChange={this._handleChange(2)}
                id="Phone" 
              />
               {this.state.IsPhoneErr == true && <span className={styles.errMsg}>{this.state.PhoneErrMsg}</span>} 
            </div>
            <div className="form-group col-md-3">
              <label className={styles.lblCtrl}>Blanket PO Request<span className={styles.star}>*</span></label>
              <input
                placeholder='Blanket PO Request'
                type="text"
                className='form-control'
                name="BlanketPORequest"
                value={this.state.BlanketPORequest}
                onChange={this._handleChange(3)}
                id="BlanketPORequest"
              />
              {this.state.IsBlanketPORequestErr == true && <span className={styles.errMsg}>{this.state.BlanketPORequestErrMsg}</span>} 
            </div>
            <div className="form-group col-md-3">
               <label className={styles.lblCtrl}>Date Required &nbsp;<span className={styles.noteMsg}>(DD-MM-YYYY)</span><span className={styles.star}>*</span></label> 
              <input
                placeholder='Date Required'
                type="date"
                className='form-control'
                name="DateRequired"
                value={this.state.DateRequired}
                onChange={this._handleChange(4)}
                id="DateRequired"
              /> 
             
              {/* <DateTimePicker 
              label="Date Required"  
          dateConvention={DateConvention.Date}  
          showLabels={false}  
         formatDate={(date: Date) => date.toLocaleDateString()}
          value={this.state.DateRequired}  
          onChange={this._handleDateChange}  
        />   */}
              {this.state.IsDateRequiredErr == true && <span className={styles.errMsg}>{this.state.DateRequiredErrMsg}</span>} 
            </div>
      

        <div className="form-group col-md-3">
              <label className={styles.lblCtrl}>Payment From<span className={styles.star}>*</span></label>
              <select className='form-control' name="PaymentForm" value={this.state.PaymentFrom} id="PaymentForm" onChange={this._handleChange(5)}>
                <option value="">Select</option>
                {this._renderDropdown(this.state.PaymentFromOptions)}
              </select>
              {this.state.IsPaymentFormErr == true && <span className={styles.errMsg}>{this.state.PaymentFormErrMsg}</span>}
        </div>

        </div>
        <div className={styles.rowTable}>
        <div className="form-group col-md-4">
              <label className={styles.lblCtrl}>Ship to Address</label><span className={styles.star}>*</span>
              <select className='form-control' name="ShipAddress" value={this.state.ShipAddress} id="ShipAddress" onChange={this._handleChange(6)}>
                <option value="">Select</option>
                {this._renderDropdown(this.state.ShipAddressOptions)}
              </select>
               {this.state.IsShipAddressErr == true && <span className={styles.errMsg}>{this.state.ShipAddressErrMsg}</span>} 
            </div>
        <div className="form-group col-md-4">
              <label className={styles.lblCtrl}>Purchasing Entity</label>
              <select className='form-control' name="PurchasingEntity" value={this.state.PurchasingEntity} id="PurchasingEntity" onChange={this._handleChange(7)}>
                <option value="">Select</option>
                {this._renderDropdown(this.state.PurchasingEntityOptions)}
              </select>
              {/* {this.state.IsPaymentFormErr == true && <span className={styles.errMsg}>{this.state.PaymentFormErrMsg}</span>} */}
            </div>
            <div className="form-group col-md-4">
              <label className={styles.lblCtrl}>Paying Entity</label>
              <select className='form-control' name="PayingEntity" value={this.state.PayingEntity} id="PayingEntity" onChange={this._handleChange(8)}>
                <option value="">Select</option>
                {this._renderDropdown(this.state.PayingEntityOptions)}
              </select>
               {/* {this.state.IsShipAddressErr == true && <span className={styles.errMsg}>{this.state.ShipAddressErrMsg}</span>}  */}
            </div>
        </div>
        {this.state.IsOtherAddress && 
        <div className="form-group col-md-12">
              <label className={styles.lblCtrl}>Other Address</label>
              <textarea
                placeholder='Other Address'
                className='form-control'
                name="OtherAddress"
                value={this.state.OtherAddress}
                onChange={this._handleChange(9)}
                id="OtherAddress">
             </textarea>

              </div>
        }
        </div>
        <table className={styles.newRequestTable}>
            <div className={this.state.IProcurementModel.length > 0?styles.groove:""}>
            {this.renderTableData()}
          </div>
            
          </table>
        <button className='btn btn-primary addItemsRow' disabled={this.state.IsBtnClicked} id="addDetailRow" onClick={this._handleAddRow}>Add New</button>&nbsp;
          {this.state.IsProcurementDetailErr == true && <span className={styles.errMsg}>{this.state.ProcurementDetailErrMsg}</span>} 
          <div>
            <div className={styles.line}></div>
          </div>
          <table className={this.state.IProcurementModel.length > 0 ? styles.newRequestTable : styles.newRequestCmtTable}>
            <tr>
              <td className={styles.cmt}>
                <label className="control-label">Comments</label>
                <textarea
                  className='form-control'
                  name="Comments"
                  value={this.state.Comments}
                  onChange={this._handleChange(5)}
                  id="Comments">
                </textarea>
              </td>
              <td className={styles.attachFile} id="inputAttachment">
                <label className="control-label">Attachment(s)</label>
                <div className={styles.noteMsg}>({this.state.AttachmentsVendorNotes})</div>
                <input className={styles.attachDoc} type="file" multiple={true} id="file" onChange={this.addFile.bind(this)} />
                {this.state.IsAttachmentErr == true && <span className={styles.errMsg}>{this.state.AttachmentErrMsg}</span>}
                </td>
               
              <td className={styles.attachedFile}>
                {this.state.fileInfos.length > 0 &&
                  <label id="fileName">Attached Files </label>
                }
                {this.renderAttachmentName()}
              </td>
            </tr>
          </table>
        <div className={ styles.btnSection}>
            <span className={styles.btnRt}>
              <button disabled={this.state.IsBtnClicked} className='btn btn-primary' id="add-row" onClick={() => this._submitRequest("Submitted")}>Submit</button>  &nbsp;&nbsp;
              <button disabled={this.state.IsBtnClicked} className='btn btn-primary' id="add-row" onClick={() => this._submitRequest("Draft")}>Save as Draft</button> &nbsp;&nbsp;
              <button disabled={this.state.IsBtnClicked} className='btn btn-primary' id="add-row" onClick={this._refreshNewRequestPage}>Cancel</button>
            </span>
          </div>
          <div id="loader" className={styles.loader}></div>
          </PivotItem>
          <PivotItem itemKey="MySubmission" key="MySubmission" headerText="My Submission" itemCount={this.state.MySubmissionItems.length} onClick={() => this.selectedTabType("MySubmission")}>
             <ListView
              items={this.state.MySubmissionItems}
              showFilter={true}
              filterPlaceHolder="Search..."
              compact={true}
              selectionMode={SelectionMode.single}
              onRenderRow={this.OnListViewRenderRow}
              listClassName={styles.listViewStyle}
              // selection={this.OpenViewFrom}    
              groupByFields={this.groupByFields()}
              viewFields={this.viewFields()}
            /> 
          </PivotItem>
           {renderPendingTabs}
          {renderApprovedTabs}
          {renderOrderPlacedTabs}
          {renderRejectedTabs} 



        </Pivot>
        {this.state.openEditDialog &&
          <Modal isOpen={this.isEditModalOpen()} isBlocking={false} containerClassName={contentEditFormStyles.container}>
            <div className={contentEditFormStyles.header}>
              <span> {this.state.selectedProcurement.formType == "Edit" ? "Edit Procurement - " : "View Procurement - "}{this.state.selectedProcurement.RequestorID}  
              
               </span>&nbsp;
               {(this.state.IsAccountant || this.state.Status=="Order Placed") && <span className={contentEditFormStyles.viewInvoice}>
               <button className='btn btn-warning' id="viewInvoice" onClick={this.OpenInvoicePopUp} title="View Invoice">View Invoice</button>
               </span>}
              <IconButton styles={iconButtonStyles} iconProps={cancelIcon} ariaLabel="Close popup modal" onClick={this.hideEditModal} />
            </div>
            {/* <EditRequestForm tabType= {this.state.SelectedTabType} selectedItem= {this.state.selectedExpense} context={this.props.context} siteUrl={this.props.siteUrl} expenseListTitle="Expenses" expenseDetailListTitle="ExpenseDetails"></EditRequestForm> */}

           <span id="generatePdfForm">
            <div className={this.state.SelectedTabType == "MyTask" || this.state.selectedProcurement.formType == "View" ? contentEditFormStyles.viewDeptSection : contentEditFormStyles.deptSection}>
        {/* <div className="row"> */}
            <div className="form-group col-md-3">
              <label className={contentEditFormStyles.lblCtrl}>Phone<span className={contentEditFormStyles.star}>*</span></label>
              <input
                placeholder='Phone'
                type="text"
                className='form-control'
                disabled={this.state.SelectedTabType == "MyTask" || this.state.selectedProcurement.formType == "View" ? true : false}
                name="Phone"
                value={this.state.Phone}
                onChange={this._handleChange(2)}
                id="Phone" 
              />
              {this.state.IsPhoneErr == true && <span className={contentEditFormStyles.errMsg}>{this.state.PhoneErrMsg}</span>} 
            </div>
            <div className="form-group col-md-3">
              <label className={contentEditFormStyles.lblCtrl}>Blanket PO Request<span className={contentEditFormStyles.star}>*</span></label>
              <input
                placeholder='Blanket PO Request'
                type="text"
                className='form-control'
                disabled={this.state.SelectedTabType == "MyTask" || this.state.selectedProcurement.formType == "View" ? true : false}
                name="BlanketPORequest"
                value={this.state.BlanketPORequest}
                onChange={this._handleChange(3)}
                id="BlanketPORequest"
              />
             {this.state.IsBlanketPORequestErr == true && <span className={contentEditFormStyles.errMsg}>{this.state.BlanketPORequestErrMsg}</span>} 
            </div>
            <div className="form-group col-md-3">
              <label className={contentEditFormStyles.lblCtrl}>Date Required &nbsp;<span className={contentEditFormStyles.noteMsg}>(DD-MM-YYYY)</span><span className={contentEditFormStyles.star}>*</span></label>
              <input
                placeholder='Date Required'
                type="date"
                data-date-format="MM-DD-YYYY"
                className='form-control'
                disabled={this.state.SelectedTabType == "MyTask" || this.state.selectedProcurement.formType == "View" ? true : false}
                name="DateRequired"
                value={this.state.DateRequired}
                onChange={this._handleChange(4)}
                id="DateRequired"
              />
             {this.state.IsDateRequiredErr == true && <span className={contentEditFormStyles.errMsg}>{this.state.DateRequiredErrMsg}</span>} 
            </div>
      

        <div className="form-group col-md-3">
              <label className={contentEditFormStyles.lblCtrl}>Payment From<span className={contentEditFormStyles.star}>*</span></label>
              <select className='form-control' disabled={this.state.SelectedTabType == "MyTask" || this.state.selectedProcurement.formType == "View" ? true : false} name="PaymentForm" value={this.state.PaymentFrom} id="PaymentForm" onChange={this._handleChange(5)}>
                <option value="">Select</option>
                {this._renderDropdown(this.state.PaymentFromOptions)}
              </select>
              {this.state.IsPaymentFormErr == true && <span className={contentEditFormStyles.errMsg}>{this.state.PaymentFormErrMsg}</span>}
            </div>
            <div className="form-group col-md-3">
              <label className={contentEditFormStyles.lblCtrl}>Ship to Address<span className={contentEditFormStyles.star}>*</span></label>
              <select className='form-control'  disabled={this.state.SelectedTabType == "MyTask" || this.state.selectedProcurement.formType == "View" ? true : false} name="ShipAddress" value={this.state.ShipAddress} id="ShipAddress" onChange={this._handleChange(6)}>
                <option value="">Select</option>
                {this._renderDropdown(this.state.ShipAddressOptions)}
              </select>
              {this.state.IsShipAddressErr == true && <span className={contentEditFormStyles.errMsg}>{this.state.ShipAddressErrMsg}</span>} 
            </div>
           
        <div className="form-group col-md-3">
              <label className={contentEditFormStyles.lblCtrl}>Purchasing Entity</label>
              <select className='form-control' disabled={this.state.SelectedTabType == "MyTask" || this.state.selectedProcurement.formType == "View" ? true : false} name="PurchasingEntity" value={this.state.PurchasingEntity} id="PurchasingEntity" onChange={this._handleChange(7)}>
                <option value="">Select</option>
                {this._renderDropdown(this.state.PurchasingEntityOptions)}
              </select>
              {/* {this.state.IsPaymentFormErr == true && <span className={contentEditFormStyles.errMsg}>{this.state.PaymentFormErrMsg}</span>} */}
            </div>
            <div className="form-group col-md-3">
              <label className={contentEditFormStyles.lblCtrl}>Paying Entity</label>
              <select className='form-control' disabled={this.state.SelectedTabType == "MyTask" || this.state.selectedProcurement.formType == "View" ? true : false} name="PayingEntity" value={this.state.PayingEntity} id="PayingEntity" onChange={this._handleChange(8)}>
                <option value="">Select</option>
                {this._renderDropdown(this.state.PayingEntityOptions)}
              </select>
               {/* {this.state.IsShipAddressErr == true && <span className={contentEditFormStyles.errMsg}>{this.state.ShipAddressErrMsg}</span>}  */}
            </div>
            <div className="form-group col-md-2">
                <label className={contentEditFormStyles.lblCtrl}>Total Amount($)</label>
                <input
                  placeholder='Total Amount'
                  type="text"
                  disabled= {true}
                  className='form-control'
                  name="TotalAmount"
                  value={this.state.TotalExpense}
                  onChange={this._handleChange(10)}
                  id="TotalAmount"
                />
              </div>
            {this.state.IsOtherAddress && 
        <div className="form-group col-md-12">
              <label className={contentEditFormStyles.lblCtrl}>Other Address</label>
              <textarea
                placeholder='Other Address'
                className='form-control'
                name="OtherAddress"
                disabled={this.state.SelectedTabType == "MyTask" || this.state.selectedProcurement.formType == "View" ? true : false}
                value={this.state.OtherAddress}
                onChange={this._handleChange(9)}
                id="OtherAddress">
             </textarea>

              </div>
        }
        {/* </div> */}


            </div>
            {/* <span className={styles.newRequestTable}>
            <div className={this.state.IProcurementModel.length > 0?styles.groove:""}>
            {this.renderTableData()}
          </div>
            
          </span> */}
             {this.renderTableDataEdit()} 
            {isShowAddProcurementBtn &&
              <div className='btn btn-primary' id="add-row"  onClick={this._handleAddRow}>Add New</div>
            }

            {/* {this.state.IsExpenseDetailErr == true && <span className={contentEditFormStyles.errMsg}>{this.state.ExpenseDetailErrMsg}</span>} */}
            <div>
              <div className={contentEditFormStyles.line}></div>
            </div>
            <div className="form-row">
              <div className="form-group col-md-4">

                <label className={contentEditFormStyles.lblCtrl}>Comments</label>
                <textarea
                  className='form-control'
                  name="LatestComments"
                  cols={6}
                  rows={3}
                  disabled={this.state.selectedProcurement.Status == "Approved" || this.state.selectedProcurement.formType == "View" ? true : false}
                  value={this.state.latestComments}
                  onChange={this._handleChange(5)}
                  id="Comments">
                </textarea>
              </div>
              {this.state.SelectedTabType != "MyTask" && this.state.selectedProcurement.formType == "Edit" &&
                <div className="form-group col-md-3">
                  <label  className={contentEditFormStyles.lblCtrl}>Attachment(s)</label>
                  <div className={contentEditFormStyles.noteMsg}>({this.state.AttachmentsVendorNotes})</div>
                  <input type="file" disabled={this.state.selectedProcurement.formType == "View" ? true : false} multiple={true} id="file" onChange={this.addFile.bind(this)} />
                  {this.state.IsAttachmentErr == true && <span className={contentEditFormStyles.errMsg}>{this.state.AttachmentErrMsg}</span>}
                </div>
              }
            
              <div className="form-group col-md-4">
                {this.state.fileInfos.length > 0 &&
                  <label id="fileName" className={contentEditFormStyles.lblCtrl}>Attached Files </label>
                }
                {this.renderAttachmentNameEdit()}

              </div>
              {this.state.selectedProcurement.formType == "View" &&
              <div className="form-group col-md-2">
              <label className={contentEditFormStyles.lblCtrl}>Submitted By</label>
                <input
                  placeholder='Submitted By'
                  type="text"
                  disabled= {true}
                  className='form-control'
                  name="Creator"
                  value={this.state.Creator}
                  onChange={this._handleChange(2)}
                  id="Creator"
                />

              </div>
  }
    {this.state.selectedProcurement.formType == "View" &&
              <div className="form-group col-md-2">
              <label className={contentEditFormStyles.lblCtrl}>Approved By</label>
                <input
                  placeholder='Approved By'
                  type="text"
                  disabled= {true}
                  className='form-control'
                  name="Manager"
                  value={this.state.ManagerStatus!=""?this.state.Manager:""}
                  onChange={this._handleChange(2)}
                  id="Manager"
                />

              </div>
  }
            </div>
            <div className="form-row">
              <div className="form-group col-md-6">
                {this.state.Comments != null &&
                  <span> <label className={contentEditFormStyles.lblCtrl}>Comments History</label>
                    <div className={contentEditFormStyles.commentContainer}>
                      <div className={contentEditFormStyles.cmtHistoryRow}>
                        <div className={contentEditFormStyles.comment}>
                          <div dangerouslySetInnerHTML={{ __html: this.state.Comments }}></div>
                        </div>
                      </div>
                    </div>
                  </span>
                }
              </div>
            </div>
            </span>
            <div>
              <span className={contentEditFormStyles.btnRt}>
                {this.state.selectedProcurement.Status != "Approved" && <span>
                  {this.state.SelectedTabType == "MySubmission" && this.state.selectedProcurement.formType == "Edit" &&
                    <span><button disabled={this.state.IsBtnClicked} className='btn btn-primary' id="add-row" onClick={() => this.UpdateRequest("Submitted")} title="Submit">Submit</button>  &nbsp;&nbsp;
                      <button disabled={this.state.IsBtnClicked} className='btn btn-primary' id="add-row" onClick={() => this.UpdateRequest("Draft")} title="Save as Draft">Save as Draft</button> &nbsp;&nbsp;
                    </span>
                  }
                  {this.state.SelectedTabType == "MyTask" && (this.state.Status!="Pending for Accountant") && (this.state.IsManager == true || this.state.IsHighManagement == true) && this.state.selectedProcurement.formType == "Edit" &&
                    <span><button disabled={this.state.IsBtnClicked} className='btn btn-primary' id="add-row" onClick={() => this.UpdateRequest("Approved")} title="Approve">Approve</button>  &nbsp;&nbsp;
                      <button disabled={this.state.IsBtnClicked} className='btn btn-primary' id="add-row" onClick={() => this.UpdateRequest("Rejected")} title="Reject">Reject</button> &nbsp;&nbsp;
                    </span>
                  }
                   {/* {this.state.SelectedTabType == "MyTask" && (this.state.ManagerId!=null?(this.state.ManagerStatus=="Approved"||this.state.ManagerStatus=="Manager Approval Not Required"):this.state.ManagerStatus=="") && this.state.Status=="Pending for Accountant" && this.state.IsAccountant == true && this.state.selectedProcurement.formType == "Edit" &&
                     <span> <button disabled={this.state.IsBtnClicked} className='btn btn-primary' id="add-row" onClick={() => this.OpenOrderPlacedPopUp()} title="Place Ordered">Place Ordered</button>&nbsp;&nbsp;
                     <button disabled={this.state.IsBtnClicked} className='btn btn-primary' id="add-row" onClick={() => this.UpdateRequest("Clarification")}>Clarification</button> &nbsp;&nbsp;
                     <button disabled={this.state.IsBtnClicked} className='btn btn-primary' id="add-row" onClick={() => this.UpdateRequest("Rejected")}>Reject</button> &nbsp;&nbsp;
                     </span>
                  }  */}
                   {this.state.SelectedTabType == "MyTask" && (this.state.selectedProcurement.Manager!=null?(this.state.ManagerStatus=="Approved"||this.state.ManagerStatus=="Manager Approval Not Required"):this.state.ManagerStatus=="") && this.state.Status=="Pending for Accountant" && this.state.IsAccountant == true && this.state.selectedProcurement.formType == "Edit" &&
                     <span> <button disabled={this.state.IsBtnClicked} className='btn btn-primary' id="add-row" onClick={() => this.OpenOrderPlacedPopUp()} title="Place Ordered">Place Ordered</button>&nbsp;&nbsp;
                     <button disabled={this.state.IsBtnClicked} className='btn btn-primary' id="add-row" onClick={() => this.UpdateRequest("Clarification")}>Clarification</button> &nbsp;&nbsp;
                     <button disabled={this.state.IsBtnClicked} className='btn btn-primary' id="add-row" onClick={() => this.UpdateRequest("Rejected")}>Reject</button> &nbsp;&nbsp;
                     </span>
                  } 
                </span>
                }

                <button disabled={this.state.IsBtnClicked} className='btn btn-primary' id="add-row" onClick={this.RefreshPage} title="Close">Close</button>
              </span>
              <div id="loaderEdit" className={contentEditFormStyles.loaderEdit}></div>
            </div>

          </Modal>
        }
           {this.state.openDialog &&
          <Modal isOpen={this.isModalOpen()} isBlocking={false} containerClassName={contentStyles.container}>
            <div className={contentStyles.body}>
              <div className={contentStyles.header}>
                <span className={styles.label}> Select Order Placed Date</span>
                <IconButton styles={iconButtonStyles} iconProps={cancelIcon} ariaLabel="Close popup modal" onClick={this.hideModal} />
              </div>

              <div className="form-row">
                <div className="form-group col-md-12">
                  <input
                    placeholder='Place Order Date'
                    type="date"
                    className='form-control'
                    name="OrderedDate"
                    value={this.state.OrderedDate}
                    onChange={this._handleChange(5)}
                    id="OrderedDate"
                  />
                  {this.state.IsOrderedDateErr == true && <span style={{ 'color': 'Red' }}>{this.state.OrderedDateErrMsg}</span>}
                </div>
              </div>

              <div className="form-row">
                <div className="form-group col-md-12">
                  <button className='btn btn-primary float-right' id="add-row" onClick={() => this.CheckOrderedDate()}>Ok</button>
                </div>
              </div>
            </div>
          </Modal>
        }
        {this.state.openInvoiceDialog &&
          <Modal isOpen={this.isInvoiceModalOpen()} isBlocking={false} containerClassName={contentInvoiceStyles.container}>
            <div>
              <div className={contentInvoiceStyles.header}>
                &nbsp;<img title="Print PDF" src={this.props.siteUrl + '/SiteAssets/ICONS/print.png'} onClick={this.documentprint} />
                <IconButton styles={iconButtonStyles} iconProps={cancelIcon} ariaLabel="Close popup modal" onClick={this.hideInvoiceModal} />
              </div>

              <div className={contentInvoiceStyles.Containerbox} id="generatePdf">
                <h1 className={contentInvoiceStyles.h1}>Procurement Invoice</h1>
                <div className="form-row">
                <div className="col-md-8">
                <div className={contentInvoiceStyles.invoiceinfo}>
                  <p><strong>Status:</strong> {this.state.Status}</p>
                  {/* </br> */}
                  <p><strong>{this.state.Creator}</strong></p>
                  <p><strong>Email:</strong> {this.state.CreatorEmail}</p>
                </div>
                  </div>
                  <div className="col-md-4">
                  <h4 className={contentInvoiceStyles.invoiceNum}><strong># {this.state.selectedProcurement.RequestorID}</strong></h4>
                  </div>
                </div>

                <div className={contentInvoiceStyles.Container}>
                  <h6 className={contentInvoiceStyles.h3}>OverAll Details</h6>
                </div>

                {/* <!-- This is OverAll Details Table  --> */}
                <table className={contentInvoiceStyles.table}>
                  <thead>
                    <tr className={contentInvoiceStyles.td}>
                      <th className={contentInvoiceStyles.th}><strong>Requestor Name</strong></th>
                      <th className={contentInvoiceStyles.th}><strong>Blanket PO Request</strong></th>
                      <th className={contentInvoiceStyles.th}><strong>Ship To Address</strong></th>
                      <th className={contentInvoiceStyles.th}><strong>Purchase Entity ($)</strong></th>
                    </tr>
                  </thead>

                  <tbody>
                    <tr>
                      <td className={contentInvoiceStyles.td}>{this.state.Creator}</td>
                      <td className={contentInvoiceStyles.td}>{this.state.BlanketPORequest}</td>
                      <td className={contentInvoiceStyles.td}>{this.state.ShipAddress}</td>

                      <td className={contentInvoiceStyles.td}>{this.state.TotalExpense}</td>
                    </tr>
                  </tbody>

                </table>
                 <br></br>

                {/* <!-- this is procurement details  --> */}

                <div className={contentInvoiceStyles.Container}>
                  <h6 className={contentInvoiceStyles.h3}>Procurement Details</h6>
                </div>
                <table className={contentInvoiceStyles.table}>
                  <thead>
                    <tr className={contentInvoiceStyles.td}>
                      <th className={contentInvoiceStyles.th}><strong>Vendors</strong></th>
                      <th className={contentInvoiceStyles.th}><strong>Item Description</strong></th>
                      <th className={contentInvoiceStyles.th}><strong>Quantity</strong></th>
                      <th className={contentInvoiceStyles.th}><strong>Unit Cost $</strong></th>
                      <th className={contentInvoiceStyles.th}><strong>$ Amount</strong></th>
                    </tr>
                  </thead>
                  {this.renderInvoiceTableData()}
                  
                  <tfoot>
                    <tr>
                      <td></td>
                      <td></td>
                      <td></td>
                      <td className={contentInvoiceStyles.total}>Total:</td>
                      <td className="total">$ {this.state.TotalExpense}</td>
                    </tr>
                  </tfoot>
                </table>
                <br></br>
                {this.state.ILogHistoryModel.length > 0 &&
                <span>
                <div className={contentInvoiceStyles.Container2}>
                  <h6 className={contentInvoiceStyles.h3}>Comment History</h6>
                </div>
                 
                {/* <!-- This is OverAll Details Table  --> */}
               
                <table className={contentInvoiceStyles.table2}>
                  <thead>
                    <tr className={contentInvoiceStyles.td}>
                      <th className={contentInvoiceStyles.th}><strong>Comment History</strong></th>
                      <th className={contentInvoiceStyles.th}><strong>Name</strong></th>
                      <th className={contentInvoiceStyles.th}><strong>Status</strong></th>
                      <th className={contentInvoiceStyles.th}><strong>Date</strong></th>
                    </tr>
                  </thead>
                  {this.renderInvoiceCommentTableData()}
                </table>
                </span>
                }
                {/* <!-- Footer Item --> */}
                <div className={contentInvoiceStyles.footeritem}>
                  <h6>Submitted By</h6>
                  <p>{this.state.Creator}</p>
                  <h6>Approved By</h6>
                  <p>{this.state.Manager}</p>
                </div>
              </div>
            </div>
          </Modal>
        }
          {this.state.openStatusBarDialog &&
          <Modal isOpen={this.isStatusBarModalOpen()} isBlocking={false} containerClassName={contentStatusBarStyles.container}>
              <div className={contentInvoiceStyles.header}>
              <span className={styles.label}> Request Id - {this.state.selectedProcurement.RequestorID}</span>
                <IconButton styles={iconButtonStyles} iconProps={cancelIcon} ariaLabel="Close popup modal" onClick={this.hideStatusBarModal} />
              </div>
              <div className={contentStatusBarStyles.progressbar} id="statuscircleId">
                  <div className={[contentStatusBarStyles.statuscircle, contentStatusBarStyles.approved].join(" ")}>
                    <i className= {`${contentStatusBarStyles.icon} fa fa-check`}></i>
                    {this.state.RequestorResponse=="Yes" && <i className={`${contentStatusBarStyles.mailIcon} fa fa-envelope`} title='Email sent' aria-hidden="true"></i>}
                    <h5 className={contentStatusBarStyles.h5}>Requestor</h5><br></br>
                    <h4 className={contentStatusBarStyles.h4}>{this.state.Status=="Draft"?"Draft":"Submitted"}</h4><br></br>
                    <h4 className={contentStatusBarStyles.h4}>({this.state.Creator})</h4>
                  </div>
                  {this.state.ManagerStatus!="Manager Approval Not Required" &&
                  <div className={managerCss}>
                    <i className={`${contentStatusBarStyles.icon} fa fa-check`}></i>
                    {this.state.ManagerResponse=="Yes" && <i className={`${contentStatusBarStyles.mailIcon} fa fa-envelope`} title='Email sent' aria-hidden="true"></i>}
                    <h5 className={contentStatusBarStyles.h5}>Manager</h5><br></br>
                    <h4 className={contentStatusBarStyles.h4}>{this.state.ManagerStatus}</h4><br></br>
                    {this.state.Manager !="" && <h4 className={contentStatusBarStyles.h4}>({this.state.Manager})</h4>}
                  </div>
          }
             {this.state.HighManagementStatus!="HighManagement Approval Not Required" && this.state.HighManagementStatus!=null && this.state.HighManagement !="" &&
                  <div className={highManagementCss}>
                    <i className={`${contentStatusBarStyles.icon} fa fa-check`}></i>
                    <div className={contentStatusBarStyles.Paid}>
                    {this.state.HighManagementResponse=="Yes" && <i className={`${contentStatusBarStyles.mailIcon} fa fa-envelope`} title='Email sent' aria-hidden="true"></i>}
                      <h5 className={contentStatusBarStyles.h5}>COE</h5><br></br>
                      <h4 className={contentStatusBarStyles.h3}>{this.state.HighManagementStatus}</h4><br></br>
                      {this.state.HighManagement !="" && <h4 className={contentStatusBarStyles.h4}>({this.state.HighManagement})</h4>}
                    </div>
                    
                  </div>
          }
                  {this.state.AccountantStatus!=null  &&
                  <div className={accountantCss}>
                    <i className={`${contentStatusBarStyles.icon} fa fa-check`}></i>
                    <div className={contentStatusBarStyles.Paid}>
                    {this.state.AccountantResponse=="Yes" && <i className={`${contentStatusBarStyles.mailIcon} fa fa-envelope`} title='Email sent' aria-hidden="true"></i>}
                      <h5 className={contentStatusBarStyles.h5}>Accountant</h5><br></br>
                      <h4 className={contentStatusBarStyles.h4}>{this.state.AccountantStatus}</h4><br></br>
                      {this.state.Accountant !="" &&  <h4 className={contentStatusBarStyles.h4}>({this.state.Accountant})</h4>}
                    </div>
                    
                  </div>
                    }
            </div>

          </Modal>
        }
      </div>
    );
  }
}
