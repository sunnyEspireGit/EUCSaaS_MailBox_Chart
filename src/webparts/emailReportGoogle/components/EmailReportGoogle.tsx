import * as React from 'react';
import styles from './EmailReportGoogle.module.scss';
import { IEmailReportGoogleProps } from './IEmailReportGoogleProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';
import * as GlobalConstants from "../../../helperFiles/constants";

export interface IEmailReportGoogleStates {
  finalArray_MS: any[];
  finalArray_Google: any[];
  totalMailCount_google : number;
  totalMailCount_ms : number;
}

export default class EmailReportGoogle extends React.Component<IEmailReportGoogleProps,IEmailReportGoogleStates, {}> {

  constructor(props) {
    super(props);
    this.state = {
      finalArray_MS: [],
      finalArray_Google: [],
      totalMailCount_google: 0,
      totalMailCount_ms:0
    };
  } 

  public render(): React.ReactElement<IEmailReportGoogleProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.emailReportGoogle} ${hasTeamsContext ? styles.teams : ''}`}>
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
                    
                    {/* <table>
                      <tr>
                        <td colSpan={18}><span style={{fontWeight:"bold"}}>Microsoft Total :</span> {this.state.totalMailCount_ms}</td>
                      </tr>
                      <tr>
                        {this.state.finalArray_MS.map(column => <><td ><span style={{fontWeight:"bold"}}>{column.SubTags} :</span> {column.occurrence}</td></>)}
                      </tr>
                    </table>

                    <br></br> */}
                    
                    <table>
                      <tr>
                        <td colSpan={8}><span style={{fontWeight:"bold"}}>Google Total :</span> {this.state.totalMailCount_google}</td>
                      </tr>
                      {/* <tr> */}
                        {this.state.finalArray_Google.map(column => <><tr><td width={"50%"}><span style={{fontWeight:"bold"}}>{column.SubTags} </span> </td><td> {column.occurrence}</td></tr></>)}
                      {/* </tr> */}
                     
                    </table>
                    <br></br>
                    {/* <table>
                      <tr>
                        
                        {this.state.finalArray.map(column => <><td ><span style={{fontWeight:"bold"}}>{column.Tag} :</span> {column.occurrence}</td></>)}
                      </tr>
                    </table> */}
                  
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

    let AllListData_Array_filter_Google = res_AllListData_Array.filter((data) => data.Tags == "Google");
    console.log("AllListData_Array_filter_Google : ", AllListData_Array_filter_Google);

    let AllListData_Array_filter_MS = res_AllListData_Array.filter((data) => data.Tags == "Microsoft");
    console.log("AllListData_Array_filter_MS : ", AllListData_Array_filter_MS);
    
  
    //Push all data into required array object
    let AllListData_Array: any[] = [];
    AllListData_Array_filter_Google.forEach((element) => {
        AllListData_Array_filter_Google.push({ ID: element.ID, text: element.Title, SubTags: element.SubTags });
    });  
    
    AllListData_Array_filter_MS.forEach((element) => {
      AllListData_Array_filter_MS.push({ ID: element.ID, text: element.Title, SubTags: element.SubTags });
    });  

    console.log("AllListData_Array_filter_Google : " , AllListData_Array_filter_Google);
    console.log("AllListData_Array_filter_MS : " , AllListData_Array_filter_MS);
            
            
    //find duplicate items count
    let finalset_google = await this.findOcc(AllListData_Array_filter_Google, "SubTags") ;
    let finalset_ms = await this.findOcc(AllListData_Array_filter_MS, "SubTags") ;
    // this.setState({finalArrayCount : finalset});
    console.log("finalset_google - ", finalset_google);
    console.log("finalset_ms - ", finalset_ms);
  
  
    this.setState({
      finalArray_Google : finalset_google, 
      finalArray_MS:finalset_ms, 
      totalMailCount_google: AllListData_Array_filter_Google.length,
      totalMailCount_ms: AllListData_Array_filter_MS.length
    });
  

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
