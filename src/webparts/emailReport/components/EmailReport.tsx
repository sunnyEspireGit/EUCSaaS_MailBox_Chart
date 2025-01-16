import * as React from 'react';
import styles from './EmailReport.module.scss';
import { IEmailReportProps } from './IEmailReportProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  sp,
  SiteGroup,
  PermissionKind,
  ICamlQuery,
  Item,
} from "@pnp/sp/presets/all";
import * as GlobalConstants from "../../../helperFiles/constants";

export interface IEmailReportStates {
  finalArray: any[];
  totalMailCount : number;
}

export default class EmailReport extends React.Component<IEmailReportProps, IEmailReportStates,{}> {

  constructor(props) {
    super(props);
    this.state = {
      finalArray: [],
      totalMailCount: 0
    };
  }  

  public render(): React.ReactElement<IEmailReportProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.emailReport} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          {/* <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        <div>
          <h3>Welcome to SharePoint Framework!</h3>
          <p>
            The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It's the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
          </p>
          <h4>Learn more about SPFx development:</h4>
          <ul className={styles.links}>
            <li><a href="https://aka.ms/spfx" target="_blank">SharePoint Framework Overview</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank">Use Microsoft Graph in your solution</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank">Build for Microsoft Teams using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank">Publish SharePoint Framework applications to the marketplace</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank">SharePoint Framework API reference</a></li>
            <li><a href="https://aka.ms/m365pnp" target="_blank">Microsoft 365 Developer Community</a></li>
          </ul> */}


          <div id='table' >
                    
                    <table>
                      <tr><td colSpan={2}><span style={{fontWeight:"bold"}}>Total :</span> {this.state.totalMailCount}</td></tr>
                     
                    </table>
                    <br></br>
                    <table>
                      <tr>
                        
                        {this.state.finalArray.map(column => <><td ><span style={{fontWeight:"bold"}}>{column.Tag} :</span> {column.occurrence}</td></>)}
                      </tr>
                    </table>
                  
            </div>
        </div>
      </section>
    );
  }


  public async componentDidMount() {
    
    this._loadAsyncData_1();
    // this.getDataFromNBHCategoryList();  
    // this.getDatafromSharePointList();   // To check users exist in which all groups
    

    // //get current logged in user details
    // await sp.web.currentUser.get().then((r) => { this.email = r.Email; this.displayName = r.Title; });
    // this.setState({ bookedFor: this.displayName });

  }

  private _loadAsyncData_1 = async() => {
    // private getDatafromSharePointList = async () => {
    // Connection to the current context's Web
    // const sp = spfi(this.context);
  
    // Get all items from List
    const res_AllListData_Array = await sp.web.lists.getByTitle(GlobalConstants.lstName_productSupport).items.select("*").getAll();
    console.log("Result : " , res_AllListData_Array);
    
  
    //Push all data into required array object
    let AllListData_Array: any[] = [];
    res_AllListData_Array.forEach((element) => {
        AllListData_Array.push({ ID: element.ID, text: element.Title, Tag:element.Tags });
    });  
    console.log("valueArray : " , AllListData_Array);
            
            
    //find duplicate items count
    let finalset = await this.findOcc(AllListData_Array, "Tag") ;
    // this.setState({finalArrayCount : finalset});
    console.log("finalset - ", finalset);
  
  
    this.setState({finalArray : finalset, totalMailCount: res_AllListData_Array.length});
  

          
  
    }

  private findOcc = async(arr, key) => {
      let arr2 = [];
  
      arr.forEach((x) => {
  
          // Checking if there is any object in arr2
          // which contains the key value
          if (arr2.some((val) => { return val[key] == x[key]; })) {
  
              // If yes! then increase the occurrence by 1
              arr2.forEach((k) => {
                  if (k[key] === x[key]) {
                      k["occurrence"]++;
                  }
              });
  
          } else {
              // If not! Then create a new object initialize 
              // it with the present iteration key's value and 
              // set the occurrence to 1
              let a = {};
              a[key] = x[key];
              a["occurrence"] = 1;
              arr2.push(a);
          }
      });
  
      return arr2;
  }
}
