import * as React from 'react';
import styles from './BsupDocGen1.module.scss';
import { IBsupDocGen1Props } from './IBsupDocGen1Props';
import { escape } from '@microsoft/sp-lodash-subset';
//import { DefaultButton } from "office-ui-fabric-react";
//import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
//import { SPHttpClient, SPHttpClientConfiguration, SPHttpClientResponse, ISPHttpClientConfiguration } from '@microsoft/sp-http';
//import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import * as Handlebars from "handlebars";
import * as HTMLDecoder from "html-decoder";
import { spfi, SPFx } from '@pnp/sp';
import { AssignFrom } from "@pnp/core";
import '@pnp/sp/webs/index.js';
import '@pnp/sp/lists';
import '@pnp/sp/items/list';
import "@pnp/sp/items/get-all";
//import { IItem, Items, PagedItemCollection } from '@pnp/sp/items';
import { IItem } from '@pnp/sp/items';

import { IItemAddResult } from "@pnp/sp/items"
import { IEmailProperties } from "@pnp/sp/sputilities";

import "@pnp/sp/sputilities";


import { DataTable, DataTableFilterMeta } from 'primereact/datatable';
import { Column } from 'primereact/column';
import { FilterMatchMode } from 'primereact/api';
import { Divider } from 'primereact/divider';
import { TabView, TabPanel } from 'primereact/tabview';
import { Tag } from 'primereact/tag';
//import { InputText } from 'primereact/inputtext'

import { Panel } from 'primereact/panel';
import 'primeflex/primeflex.css';
import "primereact/resources/primereact.min.css";    
import "primereact/resources/themes/lara-light-indigo/theme.css";   
//import "primereact/resources/themes/lara-dark-indigo/theme.css";   
import { Dialog } from 'primereact/dialog';
import { ProgressBar } from 'primereact/progressbar';
import { classNames } from 'primereact/utils';
import 'primeicons/primeicons.css';

import { PrimaryButton, Checkbox } from '@fluentui/react';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { IIconProps } from '@fluentui/react/lib/Icon';
import { TextField} from '@fluentui/react/lib/TextField';
import { ComboBox } from '@fluentui/react';
import type { IComboBoxOption} from '@fluentui/react';
import { ProgressIndicator } from '@fluentui/react/lib/ProgressIndicator';
import { MessageBar, MessageBarType} from '@fluentui/react';
import { Label } from '@fluentui/react/lib/Label';
import { Toggle } from '@fluentui/react/lib/Toggle';
//import { Link } from '@fluentui/react';
//import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';

import { CommandBarButton } from '@fluentui/react/lib/Button';
import { Text } from '@fluentui/react/lib/Text';

import { qry_documentSearch, qry_myRequests} from './CamlQuery';


        


type t_Department_Managers = {
  department: string;
  name: string;
  email: string;
}

type t_Document_Types = {
  type: string;
  class: string;
  mainDiscipline: string;
}

type MyState = {
  message: string;
  docLibData: Array<unknown>;
  myDocsLibData: Array<unknown>;
  filters: DataTableFilterMeta;
  loading: boolean;
  globalFilterValue: string;
  progressBar: boolean;
  searchError: boolean;
  sendingRequestProgress: boolean;
  sendingRequestSuccess: boolean;
  sendingRequestError: boolean;

  EquipmentType: string;
  Manufacture: string;
  Model: string;
  Asset: string;
  Site: string;
  ProposalNumber: string;  
  DocumentClass_x002d_Discpline: string;
  DocumentType: string;
  ApprovalClass: string;

  Region: string;
  Client: string;
  Serial: string;  


  requestDialogVisible: boolean;
  objRowData: {};
  userNotes: string;

  myRequestsProgress: boolean;

  documentClassOptions: IComboBoxOption[];
  documentTypeOptions: IComboBoxOption[];
  RegionOptions: IComboBoxOption[];
  ClientOptions: IComboBoxOption[];
  EquipmentTypeOptions: IComboBoxOption[];
  ManufacturerOptions: IComboBoxOption[];
  ModelOptions: IComboBoxOption[]; 
  AssetOptions: IComboBoxOption[]; 
  SiteOptions: IComboBoxOption[]; 
  SerialOptions: IComboBoxOption[];   

  PrimActiveIndex: number;

  selectedKeyRegion: string;
  selectedKeyClient: string;
  selectedKeyAsset: string;
  selectedKeySite: string;
  selectedKeyManufacturer: string;
  selectedKeyEquipmentType: string;
  selectedKeyModel: string;
  selectedKeySerial: string;
  selectedKeyDocumentlass: string;
  selectedKeyDocumentType: string;
  selectedKeyApprovalClass: string;  


  DepartmentManager: t_Department_Managers[];
  Agreed: boolean;
  AdvancedSearch: boolean;
  DocumentTypesInfo: t_Document_Types[];
  
};

/*interface SpListItem {
  Title: string;
  Id: number;
}

interface SpList {
  value: SpListItem[];
}*/

/*interface blDocItem {
  EquipmentType: string;
  Asset: string;
}*/

interface blDocRequestInfo {
  status: string;
  itemDescription: string;
}




export default class BsupDocGen1 extends React.Component<IBsupDocGen1Props, MyState> {

  state: MyState = {
    message: " ",
    docLibData: [],
    myDocsLibData: [],
    filters: {
              global: { value: null, matchMode: FilterMatchMode.CONTAINS },
              EquipmentType: { value: null, matchMode: FilterMatchMode.CONTAINS },

              Region: { value: null, matchMode: FilterMatchMode.CONTAINS },
              Client: { value: null, matchMode: FilterMatchMode.CONTAINS },
              Manufacturer: { value: null, matchMode: FilterMatchMode.CONTAINS },
              Model: { value: null, matchMode: FilterMatchMode.CONTAINS },
              Asset: { value: null, matchMode: FilterMatchMode.CONTAINS },
              Serial: { value: null, matchMode: FilterMatchMode.CONTAINS },
              DocumentClass: { value: null, matchMode: FilterMatchMode.CONTAINS },
              DocumentType: { value: null, matchMode: FilterMatchMode.CONTAINS },

              ApprovalClass: { value: null, matchMode: FilterMatchMode.CONTAINS },
              Site: { value: null, matchMode: FilterMatchMode.CONTAINS },
              ProposalNumber: { value: null, matchMode: FilterMatchMode.CONTAINS },              
            },
    loading: false,
    globalFilterValue: '',
    progressBar: false,
    searchError: false,
    sendingRequestProgress: false,
    sendingRequestSuccess: false,
    sendingRequestError: false,

    EquipmentType: '',
    Manufacture: '',
    Model: '',
    Asset: '',
    Site: '',
    ProposalNumber: '',
    DocumentClass_x002d_Discpline: '',
    DocumentType: '', 
    ApprovalClass: '',

    Region: '',
    Client: '',
    Serial: '',

    requestDialogVisible: false,
    objRowData:{},
    userNotes: '',

    myRequestsProgress: false,

    documentClassOptions: [],
    documentTypeOptions: [],
    RegionOptions: [],
    ClientOptions: [],
    EquipmentTypeOptions: [],
    ManufacturerOptions: [],    

    ModelOptions: [],
    AssetOptions: [],
    SiteOptions: [],
    SerialOptions: [],         

    PrimActiveIndex: 0,

    selectedKeyRegion: 'Blank',
    selectedKeyClient: 'Blank',
    selectedKeyAsset: 'Blank',
    selectedKeySite: 'Blank',
    selectedKeyManufacturer: 'Blank',
    selectedKeyEquipmentType: 'Blank',
    selectedKeyModel: 'Blank',
    selectedKeySerial: 'Blank',
    selectedKeyDocumentlass: 'Blank',
    selectedKeyDocumentType: 'Blank',
    selectedKeyApprovalClass: 'Blank',


    
    DepartmentManager: [], 
    Agreed: false,
    AdvancedSearch: false,
    DocumentTypesInfo: [],


  }; 

  readonly filterIcon: IIconProps = { iconName: 'Filter' };

  readonly pageSize: number = 2000;
  readonly targetDocumentList: string = 'BlackStone Document Data';  
  readonly targetRequestList: string = 'BlackStone Document Requests';

  readonly targetDocumentLibrary: string = 'BlackStone Document Library';
  
  //readonly targetDocumentSite: string = 'BlackStonePortalAdmin2';
  private get targetDocumentSite(): string {
    return this.props.targetAdminSiteWbProb?.trim() || 'BlackStoneDocPortalAdmin';
  }
  
  
  //readonly targetTenant: string = '3gxtbg';
  readonly targetViewLogsList: string = "log_Document_Views";

  readonly targetPowerAutomateSharingList: string = 'BlackStone WBP Requests';

  readonly tenantUri =  "https://" + window.location.host;
  
  readonly ApprovalClassOptions: IComboBoxOption[] = [
    { key: 'Class A', text: 'Class A' },
    { key: 'Class B', text: 'Class B' },    
    { key: 'Class C', text: 'Class C' },
    { key: '', text: '' },        
  ];  
  readonly resetFieldsIcon: IIconProps = { iconName: 'Reset' };
  
  private _resetFields = ():void => {
    // reset dropzone file picker -----------
    // reset data fields
    this.setState({Region: ''});
    this.setState({EquipmentType: ''});
    this.setState({Manufacture: ''});
    this.setState({Model: ''});
    this.setState({Asset: ''});
    this.setState({Site: ''});
    this.setState({DocumentClass_x002d_Discpline: ''});
    this.setState({DocumentType: ''});
    this.setState({ApprovalClass: ''});

    this.setState({Client: ''});
    this.setState({Serial: ''});

    this.setState({ProposalNumber: ''});

    this.setState({selectedKeyRegion: 'Blank'});
    this.setState({selectedKeyClient: 'Blank'});
    this.setState({selectedKeyAsset: 'Blank'});
    this.setState({selectedKeySite: 'Blank'});
    this.setState({selectedKeyManufacturer: 'Blank'});
    this.setState({selectedKeyEquipmentType: 'Blank'});
    this.setState({selectedKeyModel: 'Blank'});
    this.setState({selectedKeySerial: 'Blank'});
    this.setState({selectedKeyDocumentlass: 'Blank'});
    this.setState({selectedKeyDocumentType: 'Blank'});
    this.setState({selectedKeyApprovalClass: 'Blank'});  
  
  }

  async componentDidMount (): Promise<void> {
    console.log('mount start');
    await this._RefreshMyItemsClick();
    await this.loadDropDowns();
  }

  private async _populateDropDowns(index: number): Promise<void> {
    this.setState({PrimActiveIndex: index});
    if (this.state.documentClassOptions.length === 0) {      
      //await this.loadDropDowns();
    }    
    
  }

  private async loadDropDowns(): Promise<void>
  {        
    const _documentClassOptions: IComboBoxOption[] = [];
    const _documentTypeOptions: IComboBoxOption[] = [];

    const _RegionOptions: IComboBoxOption[] = [];
    const _ClientOptions: IComboBoxOption[] = [];
    const _EquipmentTypeOptions: IComboBoxOption[] = [];
    const _ManufacturerOptions: IComboBoxOption[] = [];

    const _ModelOptions: IComboBoxOption[] = [];
    const _AssetOptions: IComboBoxOption[] = [];
    const _SiteOptions: IComboBoxOption[] = [];
    const _SerialOptions: IComboBoxOption[] = [];    

    const _DepartmentManager: t_Department_Managers[] = [];      
    const _DocumentTypesInfo: t_Document_Types[] = [];
 

    // get all the items from a sharepoint list
    const sp = spfi().using(SPFx(this.props.context));
    //const spSite = spfi("https://3gxtbg.sharepoint.com/sites/" + this.targetDocumentSite).using(AssignFrom(sp.web));
    const spSite = spfi(this.tenantUri + "/sites/" + this.targetDocumentSite).using(AssignFrom(sp.web));
    
    // get all department->manager relations
    await spSite.web.lists.getByTitle("sys_Manager_Departments").items.select('Title', 'Name', 'Email').getAll().then((data) => {
      data.forEach((k) => {
        _DepartmentManager.push({department: k.Title, name: k.Name, email: k.Email});
      });      
    });
    this.setState({DepartmentManager: _DepartmentManager});


    // add request to document requests list
    await spSite.web.lists.getByTitle("sys_Document_Classes").items.select('Title').getAll().then((data) => {
      data.forEach(k => {
        _documentClassOptions.push({key:k.Title, text:k.Title});
      });      
    });
    _documentClassOptions.sort((a, b) => {
      if(a.key < b.key) { return -1; }
      if(a.key > b.key) { return 1; }
      return 0;      
    });
    _documentClassOptions.push({key:'TBD', text:'TBD'});
    _documentClassOptions.push({key:'', text:''});    
    this.setState({documentClassOptions: _documentClassOptions});

    await spSite.web.lists.getByTitle("sys_Document_Types").items.select('Title', 'ApprovalClass', 'MainDiscipline').getAll().then((data) => {
      data.forEach(k => {
        _documentTypeOptions.push({key:k.Title, text:k.Title});
        _DocumentTypesInfo.push({type: k.Title, class: k.ApprovalClass, mainDiscipline: k.MainDiscipline});
      });      
    });    
    _documentTypeOptions.sort((a, b) => {
      if(a.key < b.key) { return -1; }
      if(a.key > b.key) { return 1; }
      return 0;      
    });
    _documentTypeOptions.push({key:'TBD', text:'TBD'});
    _documentTypeOptions.push({key:'', text:''});     
    this.setState({documentTypeOptions: _documentTypeOptions});
    this.setState({DocumentTypesInfo: _DocumentTypesInfo});

    await spSite.web.lists.getByTitle("sys_Regions").items.select('Title').getAll().then((data) => {
      data.forEach(k => {
        _RegionOptions.push({key:k.Title, text:k.Title});
      });      
    });    
    _RegionOptions.sort((a, b) => {
      if(a.key < b.key) { return -1; }
      if(a.key > b.key) { return 1; }
      return 0;      
    });
    _RegionOptions.push({key:'TBD', text:'TBD'});
    _RegionOptions.push({key:'', text:''});   
    this.setState({RegionOptions: _RegionOptions});

    await spSite.web.lists.getByTitle("sys_Clients").items.select('Title').getAll().then((data) => {
      data.forEach(k => {
        _ClientOptions.push({key:k.Title, text:k.Title});
      });      
    });    
    _ClientOptions.sort((a, b) => {
      if(a.key < b.key) { return -1; }
      if(a.key > b.key) { return 1; }
      return 0;      
    });
    _ClientOptions.push({key:'TBD', text:'TBD'});
    _ClientOptions.push({key:'', text:''});   
    this.setState({ClientOptions: _ClientOptions});
    
    await spSite.web.lists.getByTitle("sys_EquipmentTypes").items.select('Title').getAll().then((data) => {
      data.forEach(k => {
        _EquipmentTypeOptions.push({key:k.Title, text:k.Title});
      });      
    });    
    _EquipmentTypeOptions.sort((a, b) => {
      if(a.key < b.key) { return -1; }
      if(a.key > b.key) { return 1; }
      return 0;      
    });
    _EquipmentTypeOptions.push({key:'TBD', text:'TBD'});
    _EquipmentTypeOptions.push({key:'', text:''}); 
    this.setState({EquipmentTypeOptions: _EquipmentTypeOptions});
    
    await spSite.web.lists.getByTitle("sys_Manufacturers").items.select('Title').getAll().then((data) => {
      data.forEach(k => {
        _ManufacturerOptions.push({key:k.Title, text:k.Title});
      });      
    });    
    _ManufacturerOptions.push({key:'TBD', text:'TBD'});
    _ManufacturerOptions.sort((a, b) => {
      if(a.key < b.key) { return -1; }
      if(a.key > b.key) { return 1; }
      return 0;      
    });
    _ManufacturerOptions.push({key:'', text:''});
    this.setState({ManufacturerOptions: _ManufacturerOptions});    

    await spSite.web.lists.getByTitle("sys_Models").items.select('Title').getAll().then((data) => {
      data.forEach(k => {
        _ModelOptions.push({key:k.Title, text:k.Title});
      });      
    });          
    _ModelOptions.sort((a, b) => {
      if(a.key < b.key) { return -1; }
      if(a.key > b.key) { return 1; }
      return 0;      
    });
    _ModelOptions.push({key:'TBD', text:'TBD'});
    _ModelOptions.push({key:'', text:''});
    this.setState({ModelOptions: _ModelOptions});    

    await spSite.web.lists.getByTitle("sys_Assets").items.select('Title').getAll().then((data) => {
      data.forEach(k => {
        _AssetOptions.push({key:k.Title, text:k.Title});
      });      
    });    
        
    _AssetOptions.sort((a, b) => {
      if(a.key < b.key) { return -1; }
      if(a.key > b.key) { return 1; }
      return 0;      
    });
    _AssetOptions.push({key:'TBD', text:'TBD'});
    _AssetOptions.push({key:'', text:''});    
    this.setState({AssetOptions: _AssetOptions});    


    await spSite.web.lists.getByTitle("sys_Sites").items.select('Title').getAll().then((data) => {
      data.forEach(k => {
        _SiteOptions.push({key:k.Title, text:k.Title});
      });      
    });    
        
    _SiteOptions.sort((a, b) => {
      if(a.key < b.key) { return -1; }
      if(a.key > b.key) { return 1; }
      return 0;      
    });
    _SiteOptions.push({key:'TBD', text:'TBD'});
    _SiteOptions.push({key:'', text:''});    
    this.setState({SiteOptions: _SiteOptions}); 

    
    await spSite.web.lists.getByTitle("sys_Serials").items.select('Title').getAll().then((data) => {
      data.forEach(k => {
        _SerialOptions.push({key:k.Title, text:k.Title});
      });      
    });              
    _SerialOptions.sort((a, b) => {
      if(a.key < b.key) { return -1; }
      if(a.key > b.key) { return 1; }
      return 0;      
    });
    _SerialOptions.push({key:'TBD', text:'TBD'});
    _SerialOptions.push({key:'', text:''});        
    this.setState({SerialOptions: _SerialOptions});            

  }  
  


  /*private Test():void {
    const myUrl = this.props.context.pageContext.web.absoluteUrl;
    this.setState({ message: myUrl})
    
  } */


  private _onGlobalFilterChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    try {
      const value = e.target.value;
      const _filters = { ...this.state.filters };

      // eslint-disable-next-line @typescript-eslint/ban-ts-comment
      // @ts-ignore
      _filters.global.value = value;

      this.setState({filters: _filters});
      //console.log("Filters: ", JSON.stringify(this.state.filters));
      this.setState({globalFilterValue: value});      
    } catch (error) {
      console.error("Filter cleared");
    }  


  };  

  private _onClearFilter = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const value = '';
    const _filters = { ...this.state.filters };
    

    // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    // @ts-ignore
    _filters.global.value = value;

    this.setState({filters: _filters});
    //console.log("Filters: ", JSON.stringify(this.state.filters));
    this.setState({globalFilterValue: value});      

  };

  private _renderHeader = (): React.ReactElement => {
    return (
        <div className="flex justify-content-end">
            <SearchBox placeholder="Filter"  iconProps={this.filterIcon} onChange={this._onGlobalFilterChange} onEscape={ev => {console.log('Custom onEscape Called');}} onClear={ev => {this._onClearFilter(ev)}}/>            
        </div>
    );
  };  

  readonly header = this._renderHeader();


  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private _actionBodyTemplate = (rowData: any): React.ReactElement => {
    //console.log('column', column);
    //console.log('rowData', JSON.stringify(rowData));    
    //return <PrimaryButton onClick={()=>this.rowColumnClick(rowData)}>Request Access</PrimaryButton>;
    return <PrimaryButton onClick={()=>this.showRequestDialog(rowData)}>Request Access</PrimaryButton>;
    
  }; 
  
  // show request dialog
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private showRequestDialog = (rowData: any): void => {
    this.setState({requestDialogVisible: true});
    this.setState({objRowData: rowData});
  }

  // save requested document's data in Requests list
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  //private rowColumnClick = async(rowData: any): Promise<void> => {
  private sendDocRequest = async(): Promise<void> => {
    
    //console.log('rowData', JSON.stringify(rowData));    
    //console.log(Object(rowData)['Asset']);
    //  console.log(Object(rowData)['EquipmentType']);
    this.setState({requestDialogVisible: false});

    this.setState({sendingRequestProgress: true});
    this.setState({sendingRequestSuccess: false});
    this.setState({sendingRequestError: false});    
    

    //reading columns of the Document data list
    const docID = Object(this.state.objRowData).docID;
    const fileName = Object(this.state.objRowData).Title;
    const equipmentType = Object(this.state.objRowData).EquipmentType;
    const manufacturer = Object(this.state.objRowData).Manufacturer;
    const model = Object(this.state.objRowData).Model;
    const asset = Object(this.state.objRowData).Asset;
    const documentClass_x002d_Discipline = Object(this.state.objRowData).DocumentClass;
    const documentType = Object(this.state.objRowData).DocumentType;
    
    //const requester = this.props.userDisplayName;
    const requester = this.props.context.pageContext.user.displayName
    const requesterEmail = this.props.context.pageContext.user.email
    const approvalClassification = Object(this.state.objRowData).ApprovalClassification;
    const accessTimeFrame = Object(this.state.objRowData).AccessTimeFrame;
    const region = Object(this.state.objRowData).Region;
    const client = Object(this.state.objRowData).Client;
    const serial = Object(this.state.objRowData).Serial;

    const site = Object(this.state.objRowData).Site;
    const proposalNumber = Object(this.state.objRowData).ProposalNumber;    
    
    let l_requestStatusTitle= "-";
    let l_emailSubject= "-";
    let l_emailBody= "-";

    try {
      const sp = spfi().using(SPFx(this.props.context));
      const spSite = spfi(this.tenantUri + "/sites/" + this.targetDocumentSite).using(AssignFrom(sp.web));

      // if Class A => then set email message and title, grant access
      // if Class B => then set email message and title and send to Admin
      // if Class C => then set email message and title and send to Manager
      
      // if documentClass_x002d_Discipline = Multi-Discipline then there will be no email mapped to this category:
      // check the sys_Document_Types list get the main discipline and use this value to lookup the manager email(s) 
      let lookUpDiscipline = '';
      lookUpDiscipline = documentClass_x002d_Discipline;
      if( lookUpDiscipline.toLowerCase().indexOf('multi') >= 0){
        console.log(this.state.DocumentTypesInfo);
        if (this.state.DocumentTypesInfo.filter((e)=> e.type === documentType).length !==0) {
          lookUpDiscipline = this.state.DocumentTypesInfo.filter((e)=> e.type === documentType)[0].mainDiscipline;
        }
        console.log('Main-Discipline-used: ' + lookUpDiscipline);
      }

      const today: string = new Date().toISOString();
      let classA_approvedDate = '';
      let email_TO =[];
      //let email_CC =[];

      
      // find document related discipline/department
      const _DepartmentManager = this.state.DepartmentManager;
      //const _requestedDepartmentManager = _DepartmentManager.filter((e)=> e.department === documentClass_x002d_Discipline);
      const _requestedDepartmentManager = _DepartmentManager.filter((e)=> e.department === lookUpDiscipline);
      const _requestedDepartmentManagerNames = _requestedDepartmentManager.map((e) => e.name);
      const _requestedDepartmentManagerEmails = _requestedDepartmentManager.map((e) => e.email);

      const _adminsEmailArray = this.props.adminEmailWbProb.split(',');
      // get the mangers emails/names ->       
      const temp_email_TO = [..._requestedDepartmentManagerEmails, ..._adminsEmailArray];
      const managerNames = _requestedDepartmentManagerNames.join(' | ');

      console.log(temp_email_TO);
      console.log(managerNames);
    
      // if Class A => then set email message and title, grant access
      if (approvalClassification === "Class A") {
        // Write to WBP list to initiate Power automate flow (sharing)
        
        const iar0: IItemAddResult = await spSite.web.lists.getByTitle(this.targetPowerAutomateSharingList).items.add({
          Title: "Grant-ClassA", //this will instruct 'Bl_SP2_ShareClassAdocsV1' flow to share the file with the email provided
          docID: Number(docID),
          EquipmentType: equipmentType,
          Manufacturer: manufacturer,
          Model: model,
          Asset:asset,
          DocumentClass: documentClass_x002d_Discipline,
          DocumentType: documentType,
          Requester: requester,
          RequesterEmail: requesterEmail,
          ApprovalClassification: approvalClassification,
          AccessTimeFrame: String(accessTimeFrame),
          Region: region,
          Client: client,
          Serial: serial,
          FileName: fileName,
          

        });
        console.log(iar0);

        l_requestStatusTitle= "Granted (Automatic)";
        l_emailSubject= "Granted Blackstone Document Request";
        l_emailBody= "Blackstone document request granted";        
      } else if (approvalClassification === "Class B") {
        l_requestStatusTitle= "Pending (Admin)";
        l_emailSubject= "Blackstone Document Request";
        l_emailBody= "Blackstone document request";          
      } else {  // Class C
        
        l_requestStatusTitle= "Pending ("  + managerNames + ")";
        l_emailSubject= "Blackstone Department Document Request";
        l_emailBody= "Blackstone document request waiting for manager approval.";        
      }



      if (approvalClassification === "Class A") { 
        classA_approvedDate = today;
        //email_TO = this.props.adminEmailWbProb; // just send email to admin
        email_TO = _adminsEmailArray;
      
      } else if (approvalClassification === "Class B") {   
        classA_approvedDate = '';
        //email_TO = this.props.adminEmailWbProb; // just send email to admin
        email_TO = _adminsEmailArray;
      } else {
        classA_approvedDate = '';
        // find document related discipline/department
        // get the mangers emails/names -> 
        //email_TO = this.props.managerEmailWbProb; // send email to manager(s) and CC to admin
        email_TO =  [..._requestedDepartmentManagerEmails, ..._adminsEmailArray];

      }

      // add request to document requests list
      const iar: IItemAddResult = await spSite.web.lists.getByTitle(this.targetRequestList).items.add({
        Title: l_requestStatusTitle,    //"Pending", Granted for Class A docs
        FileName: fileName,        
        docID: Number(docID),
        EquipmentType: equipmentType,
        Manufacturer: manufacturer,
        Model: model,
        Asset:asset,
        DocumentClass: documentClass_x002d_Discipline,
        DocumentType: documentType,
        Requester: requester,
        RequesterEmail: requesterEmail,
        ApprovalClassification: approvalClassification,
        AccessTimeFrame: String(accessTimeFrame),
        Region: region,
        Client: client,
        Serial: serial,
        Note: this.state.userNotes,
        ApprovedDate: classA_approvedDate,
      });    

      const filehref: string = this.tenantUri + '/:u:/r/sites/' + this.targetDocumentSite + '/' + encodeURIComponent(this.targetDocumentLibrary.trim()) + '/' + encodeURIComponent(fileName) + '?csf=1&web=1'

      // send email to admin
      const emailProps: IEmailProperties = {
        //To: ["saidi.ar@gmail.com"],
        To: email_TO,
        CC: [],
        BCC: [],
        Subject: l_emailSubject,
        Body: `<h2>${l_emailBody}</h2>
        <p>${requester} (${requesterEmail}) requested access to:</p>
        <p>User Note:</p>
        <p>${this.state.userNotes}</p>
            <table>
            <tbody>
              <tr>
                <td>File Name:</td>
                <td>${fileName}</td>
              </tr>
              <tr>
                <td>Region:</td>
                <td>${region}</td>
              </tr>
              <tr>
                <td>Client/Customer:</td>
                <td>${client}</td>
              </tr>
              <tr>
                <td>Equipment Type:</td>
                <td>${equipmentType}</td>
              </tr>
              <tr>
                <td>Manufacturer:</td>
                <td>${manufacturer}</td>
              </tr>
              <tr>
                <td>Model:</td>
                <td>${model}</td>
              </tr> 
              <tr>
                <td>Site:</td>
                <td>${site}</td>
              </tr>              
              <tr>
                <td>Asset:</td>
                <td>${asset}</td>
              </tr>
              <tr>
                <td>Serial #:</td>
                <td>${serial}</td>
              </tr>
              <tr>
                <td>Document Class (Discpline):</td>
                <td>${documentClass_x002d_Discipline}</td>
              </tr>
              <tr>
                <td>Document Type:</td>
                <td>${documentType}</td>
              </tr>

              <tr>
                <td>Proposal Number:</td>
                <td>${proposalNumber}</td>
              </tr>

              <tr>
                <td>Document Approval Classification:</td>
                <td>${approvalClassification}</td>
              </tr>
            </tbody>
            </table>
            <p></p>

            <a href= "${this.tenantUri + "/sites/" + this.targetDocumentSite}">Asset Management Admin site</a>
            <p></p>
            <a href = "${filehref}">Document Link</a>
            <p></p>
            <a href="https://blackstoneindustrialltd.sharepoint.com/sites/BlackstoneDocumentPortal">CAMP - Customer Asset Management Portal</a>
            <p></p>
            <p></p>`,
        AdditionalHeaders: {
            "content-type": "text/html"
        }
      };
      
      await sp.utility.sendEmail(emailProps);
      console.log("Email Sent!")

      console.log(iar);
      this.setState({sendingRequestSuccess: true});
      this.setState({sendingRequestError: false});      
      this.setState({sendingRequestProgress: false})      
    } catch (error) {
      this.setState({sendingRequestSuccess: false});
      this.setState({sendingRequestError: true});
      this.setState({sendingRequestProgress: false})      
    }

  }


  //private _getListData(): Promise<SpList> {
  /*private _getListData(): Promise<any> {
    return this.props.context.spHttpClient.get(
      this.props.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('testList')/Items", SPHttpClient.configurations.v1 )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      }).then(data => {
        //data = JSON.parse(JSON.stringify(data));
        //this.setState({ message: data.value});
        this.setState({ message: JSON.stringify(data)});
        //console.log(data.value);
      })
      .catch((e) => {
        this.setState({ message: JSON.stringify(e)})
        console.log(e);
      });  
    }*/

  /*private _getTemplateValue = (template: string, value: any): string => {
    const hTemplate = Handlebars.compile(HTMLDecoder.decode(template));
    return HTMLDecoder.decode(hTemplate(value));
  };*/    

  private _RefreshMyItemsClick = async (): Promise<void> => {
    try {
      //reset myDocsLibData array
      this.setState({ myDocsLibData: []});
      this.setState({myRequestsProgress:true});

      //const filItems = await this._searchLargeRequestList(qry_myRequests, "massoudSaidi222@3gxtbg.onmicrosoft.com");
      const filItems = await this._searchLargeRequestList(qry_myRequests, this.props.context.pageContext.user.email); // search for user records
      
  
      //this.setState({ message: filItems.length + " items found."});

      const myRequestsInfo: blDocRequestInfo[] = filItems.map(reqItem => {
        return {
          status: Object(reqItem).Title,
          fileName: Object(reqItem).FileName,
          id: Object(reqItem).ID,
          itemDescription: Object(reqItem).docID + ': ' +  Object(reqItem).EquipmentType + ' / ' + Object(reqItem).Manufacturer + ' / ' + Object(reqItem).Model + ' / ' + Object(reqItem).Asset,
          requester: Object(reqItem).Requester,
          requesterEmail: Object(reqItem).RequesterEmail,
          docClass: Object(reqItem).DocumentClass,
          approvalClass: Object(reqItem).ApprovalClassification,
          docID: Object(reqItem).docID,
          region: Object(reqItem).Region,
          client: Object(reqItem).Client,
          serial: Object(reqItem).Serial,
          equipmentType: Object(reqItem).EquipmentType,
          model: Object(reqItem).Model,
          manufacturer: Object(reqItem).Manufacturer,
          asset: Object(reqItem).Asset,
          docType: Object(reqItem).DocumentType,
          modified: Object(reqItem).Modified,
          approvedDate: Object(reqItem).ApprovedDate,
          accessTimeFrame: Object(reqItem).AccessTimeFrame,
          created: Object(reqItem).Created,

        };
      });      
      
      //await this.loadDropDowns();

      //this.setState({ myDocsLibData: filItems});
      this.setState({ myDocsLibData: myRequestsInfo});
      this.setState({myRequestsProgress:false});    
      //this.setState({ searchError  : false});
    } catch (error) {
      //this.setState({ message: " ERROR"});
      this.setState({myRequestsProgress:false});    
      //this.setState({ searchError  : true});          
    }
  }

  private _allItemsClick = async (): Promise<void> => {

    this.setState({ progressBar: true});
    this.setState({ searchError  : false}); 
    //const sp = spfi().using(SPFx(this.props.context));
    //const filItems = await sp.web.lists.getByTitle('BlackStone Document Library').items.getAll();
    //const filItems = await sp.web.lists.getByTitle(this.targetDocumentList).items.getAll();    

    console.log("this.state.Region: " + this.state.Region);

    if (this.state.Region === "" && this.state.Client === "" && this.state.Serial === "" && this.state.EquipmentType === "" && this.state.Manufacture === "" && this.state.Model === "" && this.state.Asset === "" && this.state.DocumentClass_x002d_Discpline === "" && this.state.DocumentType  === "" && this.state.Site === "" && this.state.ProposalNumber  === "") {
      this.setState({ message: "Please specify at least one search criteria and try again."});
      this.setState({ progressBar: false});  
      this.setState({ searchError  : false});
    } else {
      try {
        const filItems = await this._searchLargeList(qry_documentSearch, this.state.Region , this.state.Client , this.state.Serial, this.state.EquipmentType , this.state.Manufacture , this.state.Model, this.state.Asset, this.state.DocumentClass_x002d_Discpline, this.state.DocumentType, this.state.Site, this.state.ProposalNumber);
        //console.log("equipment:...." + this.state.EquipmentType);
        //console.log("Normalized Caml Query filter with EquipmentType: ", filItems);
    
        //console.log("Get All Items: ", filItems);
        this.setState({ message: filItems.length + " item(s) found."});
        this.setState({ docLibData: filItems});
        this.setState({ progressBar: false});    
        this.setState({ searchError  : false});
      } catch (error) {
        this.setState({ message: " ERROR"});
        this.setState({ progressBar: false});  
        this.setState({ searchError  : true});          
      }
    }



  }

  private _getMaxIdForList = async (listname: string): Promise<number> => {
    const sp = spfi().using(SPFx(this.props.context));
    const spSite = spfi(this.tenantUri + "/sites/" + this.targetDocumentSite).using(AssignFrom(sp.web));
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const maxItems: any[] = await spSite.web.lists.getByTitle(listname).items
        .select('ID')
        .orderBy('ID', false)
        .top(1)();
    if (maxItems.length > 0) return maxItems[0].ID;
    else return 0;
  };  

  private _searchLargeRequestList = async (itemsQuery: string, RequesterSearchKey: string): Promise<IItem[]> => {
    let minid: number;
    let maxid: number;

    const targetLibOrListName = this.targetRequestList

    const listmaxid: number = await this._getMaxIdForList(targetLibOrListName);
    const maxPage: number = Math.ceil(listmaxid / this.pageSize);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let returnItems: any[] = [];
    const sp = spfi().using(SPFx(this.props.context));
    const spSite = spfi(this.tenantUri + "/sites/" + this.targetDocumentSite).using(AssignFrom(sp.web));

    for (let i = 0; i < maxPage; i++) {
        minid = i * this.pageSize + 1;
        maxid = (i + 1) * this.pageSize;
        console.log(`Min id: ${minid.toString()} - Max Id: ${maxid.toString()}`);
        const camlQuery: string = this._getTemplateValue(itemsQuery, { 
          RequesterSearchKey: RequesterSearchKey, 
          minid: minid, maxid: maxid });
        const retitems: IItem[] = await spSite.web.lists.getByTitle(targetLibOrListName).getItemsByCAMLQuery({
            ViewXml: camlQuery
        });
        if (retitems.length > 0) {
            returnItems = returnItems.concat(retitems);
        }
        if (i >= maxPage - 1) return returnItems;

    }

    return returnItems;
  };



  private _searchLargeList = async (itemsQuery: string, RegionSearchKey: string, ClientSearchKey: string, SerialSearchKey: string, EquipmentTypeSearchKey: string, ManufactureSearchKey: string, ModelSearchKey: string, AssetSearchKey: string, DocumentClass_DiscplineSearchKey: string, DocumentTypeSearchKey: string, SiteSearchKey: string, ProposalNumberSearchKey: string): Promise<IItem[]> => {
    let minid: number;
    let maxid: number;
    const listmaxid: number = await this._getMaxIdForList(this.targetDocumentList);
    const maxPage: number = Math.ceil(listmaxid / this.pageSize);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let returnItems: any[] = [];
    const sp = spfi().using(SPFx(this.props.context));
    const spSite = spfi(this.tenantUri + "/sites/" + this.targetDocumentSite).using(AssignFrom(sp.web));

    //let EquipmentTypeFieldName: string;
    //(EquipmentTypeSearchKey) ? (EquipmentTypeFieldName = "EquipmentType") : (EquipmentTypeFieldName = "SystemNull");
    //const Contains_EquipmentTypeSearchKey: string = 'Contains';
    const Contains_RegionSearchKey = (RegionSearchKey !== '[Any]' && RegionSearchKey !== '') ? 'Contains': 'Neq';
    const Contains_ClientSearchKey = (ClientSearchKey !== '[Any]' && ClientSearchKey !== '') ? 'Contains': 'Neq';
    const Contains_SerialSearchKey = SerialSearchKey !== '' ? 'Contains': 'Neq';
    const Contains_EquipmentTypeSearchKey = (EquipmentTypeSearchKey !== '[Any]' && EquipmentTypeSearchKey !== '') ? 'Contains': 'Neq';
    const Contains_ManufactureSearchKey = (ManufactureSearchKey !== 'Any' && ManufactureSearchKey !== '')? 'Contains': 'Neq';
    const Contains_ModelSearchKey = ModelSearchKey !== '' ? 'Contains': 'Neq';
    const Contains_AssetSearchKey = AssetSearchKey !== '' ? 'Contains': 'Neq';
    const Contains_DocumentClass_DiscplineSearchKey = (DocumentClass_DiscplineSearchKey !== 'Any' && DocumentClass_DiscplineSearchKey !== '') ? 'Contains': 'Neq';
    const Contains_DocumentTypeSearchKey = (DocumentTypeSearchKey !== 'Any' && DocumentTypeSearchKey !== '') ? 'Contains': 'Neq';    

    const Contains_SiteSearchKey = (SiteSearchKey !== '[Any]' && SiteSearchKey !== '') ? 'Eq': 'Neq';
    const Contains_ProposalNumberSearchKey = (ProposalNumberSearchKey !== '[Any]' && ProposalNumberSearchKey !== '') ? 'Eq': 'Neq';    

    for (let i = 0; i < maxPage; i++) {
        minid = i * this.pageSize + 1;
        maxid = (i + 1) * this.pageSize;
        console.log(`Min id: ${minid.toString()} - Max Id: ${maxid.toString()}`);
        const camlQuery: string = this._getTemplateValue(itemsQuery, { 
          RegionSearchKey: RegionSearchKey,
          ClientSearchKey: ClientSearchKey,
          SerialSearchKey: SerialSearchKey,
          
          EquipmentTypeSearchKey: EquipmentTypeSearchKey, 
          ManufactureSearchKey: ManufactureSearchKey, 
          ModelSearchKey: ModelSearchKey, 
          AssetSearchKey: AssetSearchKey,
          DocumentClass_DiscplineSearchKey: DocumentClass_DiscplineSearchKey,
          DocumentTypeSearchKey: DocumentTypeSearchKey, 

          SiteSearchKey: SiteSearchKey,
          ProposalNumberSearchKey: ProposalNumberSearchKey,          
          
          Contains_RegionSearchKey: Contains_RegionSearchKey,
          Contains_ClientSearchKey: Contains_ClientSearchKey,
          Contains_SerialSearchKey: Contains_SerialSearchKey,

          Contains_EquipmentTypeSearchKey: Contains_EquipmentTypeSearchKey,
          Contains_ManufactureSearchKey : Contains_ManufactureSearchKey,
          Contains_ModelSearchKey: Contains_ModelSearchKey,
          Contains_AssetSearchKey : Contains_AssetSearchKey,
          Contains_DocumentClass_DiscplineSearchKey : Contains_DocumentClass_DiscplineSearchKey, 
          Contains_DocumentTypeSearchKey : Contains_DocumentTypeSearchKey ,

          Contains_SiteSearchKey: Contains_SiteSearchKey,
          Contains_ProposalNumberSearchKey: Contains_ProposalNumberSearchKey,          

          minid: minid, maxid: maxid });
        
          //console.log(camlQuery);
        const retitems: IItem[] = await spSite.web.lists.getByTitle(this.targetDocumentList).getItemsByCAMLQuery({
            ViewXml: camlQuery
        });
        if (retitems.length > 0) {
            returnItems = returnItems.concat(retitems); 

        }
        if (i >= maxPage - 1) return returnItems;
    }
    return returnItems;
  };


  private _getTemplateValue = (template: string, value: unknown): string => {
    const hTemplate = Handlebars.compile(HTMLDecoder.decode(template));
    return HTMLDecoder.decode(hTemplate(value));
  };
 

  /*private _updateInputEquipmentTypeValue(evt: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>): void {
    const val = evt.currentTarget.value;
    console.log(val);
    this.setState({EquipmentType: val});    
    
  }*/

  private _dismissRequestMessages = (): void => {
    this.setState({sendingRequestSuccess: false});
    this.setState({sendingRequestError: false});
    this.setState({sendingRequestProgress: false});
  }

  
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private requestStatusBodyTemplate = (rowData: any): React.ReactElement => {
    return <Tag value={rowData.status} severity={this.getRequestStatus(rowData)}/>;
  };

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private approvedDateBodyTemplate = (rowData: any): React.ReactElement => {
    let days = 0;
    const oneDay = 1000 * 60 * 60 * 24;
    try {
      if (String(rowData.status).indexOf("Granted") >= 0) {
        days = (new Date()).valueOf() - (new Date(rowData.approvedDate)).valueOf();  
        const diffInDays = Math.round(days / oneDay);
        const timeFrame = Number(rowData.accessTimeFrame)
        //console.log("diffInday: " + diffInDays);
        //console.log("timeFrame: " + timeFrame);
        if (diffInDays <= timeFrame) {
          const prog = (diffInDays / timeFrame ) * 100;
          return <ProgressBar value={prog} showValue={false} style={{ height: '6px' }}/>;
        } else {
          return <i className={classNames('pi', 'text-red-500 pi-times-circle')}/>;
        }
        
      } else if (String(rowData.status).indexOf("Unshared") >= 0) {
        return <i className={classNames('pi', 'text-red-500 pi-ban')}/>;
      } else {  // pending or other status
        //return <i className={classNames('pi', { 'text-green-500 pi-check-circle': rowData.verified, 'text-red-500 pi-times-circle': !rowData.verified })}/>;
        return <i className={classNames('pi', 'text-red-500 pi-ellipsis-h')}/>;
        
      }
    } catch (error) {      
      return <i className={classNames('pi', 'text-red-500 pi-exclamation-triangle')}/>;
    }
    
    
  };  

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private linkBodyTemplate = (rowData: any): React.ReactElement => {
    //to force open a document in the web browser
    //Csf=1 – added when sharing with People with existing access
    //Web=1 – opens the file in the browser
    //https://cresyl.eu/en/office-365/understanding-how-share-links-are-built/
    //#Read more: https://www.sharepointdiary.com/2020/12/open-document-in-browser-instead-of-client-application-in-sharepoint-online.html#ixzz8AuvsX531
    //https://3gxtbg.sharepoint.com/:u:/r/sites/BlackStonePortalAdmin/BlackStone%20Document%20Library/debug.log?csf=1&web=1&e=PLccIF

    let days = 0;
    const oneDay = 1000 * 60 * 60 * 24;
    if (String(rowData.status).indexOf('Granted') >= 0) {    
      try {
        days = (new Date()).valueOf() - (new Date(rowData.approvedDate)).valueOf();  
        const diffInDays = Math.round(days / oneDay);
        const timeFrame = Number(rowData.accessTimeFrame)

        if (diffInDays <= timeFrame) {          
          //return <a href="#" onClick={()=> this.logAndShowFiles(rowData)}>Link</a>
          return <button className={styles.buttonFourteen} role="button" onClick={()=> this.logAndShowFiles(rowData)}>view</button>;
        } else {
          return ;
        }
      } catch (error) {      
        return ;
      }        


    } else {
      return ;
    }
  };

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private logAndShowFiles = async (rowData: any): Promise<void> => {
    //const filehref: string = 'https://' + this.targetTenant + '.sharepoint.com/:u:/r/sites/' + this.targetDocumentSite + '/' + encodeURIComponent(this.targetDocumentLibrary.trim()) + '/' + encodeURIComponent(rowData.fileName) + '?csf=1&web=1'
    const filehref: string = this.tenantUri + '/:u:/r/sites/' + this.targetDocumentSite + '/' + encodeURIComponent(this.targetDocumentLibrary.trim()) + '/' + encodeURIComponent(rowData.fileName) + '?csf=1&web=1'
    const myUrl: URL = new URL(filehref);
    
    // log file views in targetViewLogsList
    const sp = spfi().using(SPFx(this.props.context));
    const spSite = spfi(this.tenantUri + "/sites/" + this.targetDocumentSite).using(AssignFrom(sp.web));

    // add request to document requests list
    await spSite.web.lists.getByTitle(this.targetViewLogsList).items.add({
      Title: "view",
      FileName: Object(rowData).fileName,        
      docID: Object(rowData).docID,
      EquipmentType: Object(rowData).equipmentType,
      Manufacturer: Object(rowData).manufacturer,
      Model: Object(rowData).model,
      Asset:Object(rowData).asset,
      DocumentClass: Object(rowData).docClass,
      DocumentType: Object(rowData).docType,
      Requester: this.props.context.pageContext.user.displayName,
      RequesterEmail: this.props.context.pageContext.user.email,
      ApprovalClassification: Object(rowData).approvalClass,
      AccessTimeFrame: Object(rowData).accessTimeFrame,
      Region: Object(rowData).region,
      Client: Object(rowData).client,
      Serial: Object(rowData).serial,
    });        

    // open in new tab
    window.open(myUrl, '_blank', 'noreferrer')
  }

  
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private getRequestStatus = (rowData: any): any => {

    if (String(rowData.status).indexOf('Granted') >= 0) {
      return 'success';
    } else if (String(rowData.status).indexOf('Pending') >= 0) {
      return 'warning';
    } else if (String(rowData.status).indexOf('Unshared') >= 0) {
      return 'danger';
    } else {
      return null;
    }

    /*switch (rowData.status) {
      case 'Granted':
        return 'success';

      case 'Pending':
          return 'warning';

      case 'Unshared':
          return 'danger';

      default:
          return null;
    }*/
  };  

  private onBsupProposalNumberChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    this.setState({ProposalNumber: newValue});
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private onBsupComboChange = (selectedOption: any, value: string, bsupField: string): void => {
    const v = value;
    const key = selectedOption?.key;
    //if (selectedOption) {
    //    v = selectedOption.key;
    //}
    //this.props.onStateChange(v);
    if (bsupField === 'Region') {
      this.setState({Region: v});
      this.setState({selectedKeyRegion: key});
    }
    if (bsupField === 'Client') {
      this.setState({Client: v});
      this.setState({selectedKeyClient: key});
    }
    if (bsupField === 'EquipmentType') {
      this.setState({EquipmentType: v});
      this.setState({selectedKeyEquipmentType: key});
    }
    if (bsupField === 'Manufacturer') {
      this.setState({Manufacture: v});
      this.setState({selectedKeyManufacturer: key});
    }
    if (bsupField === 'documentClass') {
      this.setState({DocumentClass_x002d_Discpline: v});
      this.setState({selectedKeyDocumentlass: key});
    }
    if (bsupField === 'documentType') {
      this.setState({DocumentType: v});
      this.setState({selectedKeyDocumentType: key}); 
    }    
    if (bsupField === "ApprovalClass") {
      this.setState({ApprovalClass: v});
      this.setState({selectedKeyApprovalClass: key});
    }

    if (bsupField === "model") {
      this.setState({Model: v});
      this.setState({selectedKeyModel: key});
    }
    if (bsupField === "asset") {
      this.setState({Asset: v});
      this.setState({selectedKeyAsset: key});
    }

    if (bsupField === "Site") {
      this.setState({Site: v});
      this.setState({selectedKeySite: key});
    }

    if (bsupField === "serial") {
      this.setState({Serial: v});
      this.setState({selectedKeySerial: key});
    }
  } 

  private _onAgreementChange = (ev?: React.FormEvent<HTMLElement | HTMLInputElement>, isChecked?: boolean): void => {
    if (isChecked) {
      this.setState({Agreed: true}); 
    }
    else {
      this.setState({Agreed: false}); 
    }
  }  

  private _onAdvancedSearchChange = (ev: React.MouseEvent<HTMLElement>, checked?: boolean): void => {
    if (checked) {
      this.setState({AdvancedSearch: true}); 
    }
    else {
      this.setState({AdvancedSearch: false}); 
    }
  }

  public render(): React.ReactElement<IBsupDocGen1Props> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,      
    } = this.props;

    return (
      <section className={`${styles.bsupDocGen1} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <div className={styles.mas_left + ' ' + styles.mas_clear}>
            <img alt="" src={isDarkTheme ? require('../assets/BlackStone-logo-256V2-dark.png') : require('../assets/BlackStone-logo-256V2-light.png')} className={styles.LogoImage} />
          </div>
          <div className={styles.mas_right}><h1>Document Portal</h1></div>          
        </div> 
        <Divider />       
        <div className={styles.welcome}>          
          <div>Welcome, {escape(userDisplayName)}</div>
          <div>{environmentMessage}</div>
          <div>Version: <strong>3.6</strong>  {escape(description)}</div>
        </div>
        <div>
        <Divider />  
      
        <TabView activeIndex={this.state.PrimActiveIndex} onTabChange={(e)=>this._populateDropDowns(e.index)}>
          <TabPanel header="My Files">
            <div className="flex flex-wrap gap-3 mb-4">
                <div className="flex-auto">
                  <PrimaryButton iconProps={{ iconName: "refresh" }} onClick={()=>this._RefreshMyItemsClick()}>
                    Refresh
                  </PrimaryButton>                  
                </div>
                <div className="flex-auto">
                  {this.state.myRequestsProgress && <ProgressIndicator label="Loading..." description=""/>}
                </div>
            </div>

            <DataTable value={this.state.myDocsLibData} sortField="price" sortOrder={-1} scrollable emptyMessage="No items found." showGridlines stripedRows size="small" loading={this.state.loading}>
                <Column field="id" header="ID" sortable/>                
                <Column header="Status" body={(rowData: unknown) => this.requestStatusBodyTemplate(rowData)} style={{ minWidth: '14rem' }}/>                                
                <Column header="Link" body={(rowData: unknown) => this.linkBodyTemplate(rowData)}/>                                
                                        
                <Column field="fileName" header="File Name" sortable style={{ minWidth: '12rem' }}/>                
                <Column field="approvedDate" header="Access" body={this.approvedDateBodyTemplate} style={{ minWidth: '4rem' }}/>                
                <Column field="accessTimeFrame" header="Frame" style={{ minWidth: '2rem' }}/>                
                <Column field="approvedDate" header="Approved Date"  style={{ minWidth: '12rem' }}/> 
                <Column field="created" header="Registered Date" style={{ minWidth: '12rem' }}/>
                <Column field="docClass" header="Doc. Class" sortable style={{ minWidth: '12rem' }}/>
                <Column field="approvalClass" header="Approval" sortable style={{ minWidth: '6rem' }}/>
                <Column field="itemDescription" header="Description" style={{ minWidth: '24rem' }}/>                                                         
            </DataTable>

          </TabPanel>
          <TabPanel header="Search Documents">
              
              <Panel header="" toggleable>
                
                <div className="formgrid grid">

                <div className="field col-12 md:col-12">
                    <CommandBarButton iconProps={this.resetFieldsIcon} text="Reset Fields" onClick={()=>this._resetFields()}/>
                </div>

                <div className="field col-12 md:col-4">
                    <ComboBox
                        label="Region"
                        options={this.state.RegionOptions}
                        autoComplete="on"
                        onChange={(e, selectedOption, i, value) => { this.onBsupComboChange(selectedOption, value, "Region"); }}  
                        selectedKey={this.state.selectedKeyRegion}
                      />                     
                  </div>
                  <div className="field col-12 md:col-4">
                    <ComboBox
                        label="Client/Customer"
                        options={this.state.ClientOptions}
                        autoComplete="on"
                        onChange={(e, selectedOption, i, value) => { this.onBsupComboChange(selectedOption, value, "Client"); }}  
                        selectedKey={this.state.selectedKeyClient}
                      />                                         
                  </div>

                  <div className="field col-12 md:col-4">
                    <ComboBox
                        label="Site"
                        options={this.state.SiteOptions}
                        autoComplete="on"
                        onChange={(e, selectedOption, i, value) => { this.onBsupComboChange(selectedOption, value, "Site"); }}  
                        selectedKey={this.state.selectedKeySite}
                      />                                         
                  </div>                    

                  <div className="field col-12 md:col-4">
                      <ComboBox
                        label="Asset"
                        options={this.state.AssetOptions}
                        autoComplete="on"
                        onChange={(e, selectedOption, i, value) => { this.onBsupComboChange(selectedOption, value, "asset"); }}  
                        selectedKey={this.state.selectedKeyAsset}
                      />                    
                  </div>

                  <div className="field col-12 md:col-4">
                      <ComboBox
                        label="Serial #"
                        options={this.state.SerialOptions}
                        autoComplete="on"
                        onChange={(e, selectedOption, i, value) => { this.onBsupComboChange(selectedOption, value, "serial"); }}  
                        selectedKey={this.state.selectedKeySerial}
                      />                    
                  </div>                  

                  <div className="field col-12 md:col-4">
                    {this.state.AdvancedSearch && <Toggle
                      label={<div>Advanced search options</div>}
                      defaultChecked
                      onText="On"
                      offText="Off"
                      onChange={this._onAdvancedSearchChange}
                    />}
                    {!this.state.AdvancedSearch && <Toggle
                      label={<div>Advanced search options</div>}
                      onText="On"
                      offText="Off"
                      onChange={this._onAdvancedSearchChange}
                    />}                    
                  </div>

                  {this.state.AdvancedSearch && <div className="field col-12 md:col-4">
                    <ComboBox
                        label="Equipment Type"
                        options={this.state.EquipmentTypeOptions}
                        autoComplete="on"
                        onChange={(e, selectedOption, i, value) => { this.onBsupComboChange(selectedOption, value, "EquipmentType"); }}  
                        selectedKey={this.state.selectedKeyEquipmentType}
                      />                                                             
                  </div>}




                  {this.state.AdvancedSearch && <div className="field col-12 md:col-4">
                    <ComboBox
                      label="Model"
                      options={this.state.ModelOptions}
                      autoComplete="on"
                      onChange={(e, selectedOption, i, value) => { this.onBsupComboChange(selectedOption, value, "model"); }}  
                      selectedKey={this.state.selectedKeyModel}
                    />                    
                  </div>}



                  {this.state.AdvancedSearch && <div className="field col-12 md:col-4">
                      <ComboBox
                        label="Manufacturer"
                        options={this.state.ManufacturerOptions}
                        autoComplete="on"
                        onChange={(e, selectedOption, i, value) => { this.onBsupComboChange(selectedOption, value, "Manufacturer"); }}  
                        selectedKey={this.state.selectedKeyManufacturer}
                      />                                                                                   
                  </div>}

                  {this.state.AdvancedSearch && <div className="field col-12 md:col-4">
                    <ComboBox
                      label="Document Class-Discipline"
                      options={this.state.documentClassOptions}
                      autoComplete="on"
                      onChange={(e, selectedOption, i, value) => { this.onBsupComboChange(selectedOption, value, "documentClass"); }}  
                      selectedKey={this.state.selectedKeyDocumentlass}
                    />                    
                  </div>}

                  {this.state.AdvancedSearch && <div className="field col-12 md:col-4">
                    <ComboBox
                      label="Document Type"
                      options={this.state.documentTypeOptions}
                      autoComplete="on"
                      onChange={(e, selectedOption, i, value) => { this.onBsupComboChange(selectedOption, value, "documentType"); }}  
                      selectedKey={this.state.selectedKeyDocumentType}
                    /> 
                  </div>}

                  {this.state.AdvancedSearch && <div className="field col-12 md:col-4">
                      <TextField 
                        label="Proposal Number" 
                        value={this.state.ProposalNumber}
                        onChange={this.onBsupProposalNumberChange}    
                      />
                  </div>}


                  {this.state.AdvancedSearch && <div className="field col-12 md:col-4">
                    <Text variant={'small'} block>Document Approval Class Definitions:</Text>
                    <Text variant={'small'} block>Class A - View only with immediate (~ 1 min.) but monitored access.</Text>
                    <Text variant={'small'} block>Class B - View only with access granted by Document Control.</Text>
                    <Text variant={'small'} block>Class C - View only with access granted by Management.</Text>
                  </div>}

                  <div className="field col-12 md:col-4">
                    <PrimaryButton iconProps={{ iconName: "Search" }} onClick={()=>this._allItemsClick()}>
                      Search Documents
                    </PrimaryButton>
                  </div>

                  <div className="field col-12 md:col-8">
                    {!this.state.progressBar && <Label>{this.state.message}</Label>}
                  </div>
                </div>

                {this.state.progressBar && <ProgressIndicator label="Loading server records..." description=""/>}                
                {this.state.searchError && <MessageBar
                      delayedRender={false}
                      // Setting this to error, blocked, or severeWarning automatically sets the role to "alert"
                      messageBarType={MessageBarType.error}
                      // Or you could set the role manually, IF an alert role is appropriate for the message
                      // role="alert"
                      >
                        Error in search or connection.
                      </MessageBar>}            
              </Panel>

              <br/>
      
              {this.state.sendingRequestError && <MessageBar messageBarType={MessageBarType.error} onDismiss={this._dismissRequestMessages}>Error in sending document request.</MessageBar>}
              {this.state.sendingRequestSuccess && <MessageBar messageBarType={MessageBarType.success} onDismiss={this._dismissRequestMessages}>Your document request was sent successfully.</MessageBar>}
              {this.state.sendingRequestProgress && <ProgressIndicator label="Please wait sending your request..." description=""/>}          
              <div className="card">
                <DataTable value={this.state.docLibData} scrollable paginator rows={10}  rowsPerPageOptions={[5, 10, 25, 50]} header={this.header} emptyMessage="No items found." showGridlines stripedRows size="small" removableSort 
                        dataKey="id" filters={this.state.filters} filterDisplay="row" loading={this.state.loading}
                        globalFilterFields={['Region', 'Client','EquipmentType','Manufacturer','Model','Asset', 'Serial','DocumentClass','DocumentType', 'ApprovalClassification', 'Site', 'ProposalNumber']}>
                  <Column field="Action" header="Action"  body={(rowData: unknown) => this._actionBodyTemplate(rowData)}/>
                  <Column field="Title" header="File Name" sortable style={{ minWidth: '12rem' }} />
                  <Column field="ApprovalClassification" header="Approval Class" sortable style={{ minWidth: '12rem' }}/>
                  <Column field="Region" header="Region" sortable filter filterPlaceholder="Region" style={{ minWidth: '12rem' }} />
                  <Column field="Client" header="Client" sortable filter filterPlaceholder="Client" style={{ minWidth: '12rem' }} />                  
                  <Column field="EquipmentType" header="EquipmentType" sortable filter filterPlaceholder="EquipmentType" style={{ minWidth: '12rem' }} />
                  <Column field="Manufacturer" header="Manufacturer" sortable filter filterPlaceholder="Manufacturer" style={{ minWidth: '12rem' }}/>
                  <Column field="Model" header="Model" sortable filter filterPlaceholder="Model" style={{ minWidth: '12rem' }}/>
                  <Column field="Site" header="Site" sortable filter filterPlaceholder="Site" style={{ minWidth: '12rem' }}/>                  
                  <Column field="Asset" header="Asset" sortable filter filterPlaceholder="Asset" style={{ minWidth: '12rem' }}/>
                  <Column field="Serial" header="Serial" sortable filter filterPlaceholder="Serial" style={{ minWidth: '12rem' }} />                  
                  <Column field="DocumentClass" header="DocumentClass-Discipline" sortable filter filterPlaceholder="Class" style={{ minWidth: '12rem' }}/>
                  <Column field="DocumentType" header="Document Type" sortable filter filterPlaceholder="Type" style={{ minWidth: '12rem' }}/>                  
                  <Column field="ProposalNumber" header="Proposal #" sortable filter filterPlaceholder="Proposal #" style={{ minWidth: '12rem' }}/>                  
                </DataTable>
              </div>  
            </TabPanel>
          </TabView>        
        </div>



        <Dialog header="Optional Note for Admin" visible={this.state.requestDialogVisible} style={{ width: '50vw' }} onHide={() => this.setState({requestDialogVisible: false})} footer={this.state.Agreed && <div><PrimaryButton iconProps={{ iconName: "Send" }} onClick={()=>this.sendDocRequest()}>Send Request</PrimaryButton></div>}>
          <TextField label="" borderless multiline placeholder="Enter your (optional) note for Admin here." onChange={evt => this.setState({userNotes: evt.currentTarget.value})}/>
          <Label>Please note:</Label>
          <Label>This document is confidential and according to BLACKSTONE’s agreement, this document shall not be shared, printed, downloaded, screenshot or reproduced in any way. By clicking OK to this request means you agree and acknowledge . If you do not agree, you will not be able to request to access the document</Label>
          <Checkbox label="I confirm and acknowledge." checked={this.state.Agreed} onChange={this._onAgreementChange} />
          
        </Dialog>        

      </section>
      
    );
  }
}
