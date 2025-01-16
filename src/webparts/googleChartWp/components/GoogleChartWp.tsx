import * as React from "react";
import styles from "./GoogleChartWp.module.scss";
import { IGoogleChartWpProps } from "./IGoogleChartWpProps";
import { escape } from "@microsoft/sp-lodash-subset";
import {
  ChartControl,
  ChartPalette,
  ChartType,
} from "@pnp/spfx-controls-react/lib/ChartControl";
// import { Chart } from 'chart.js';
import {
  sp,
  SiteGroup,
  PermissionKind,
  ICamlQuery,
  Item,
} from "@pnp/sp/presets/all";
// import { getSP } from "../pnpjsConfig";
import { Spinner, SpinnerSize } from "office-ui-fabric-react/lib/Spinner";
import {
  Dropdown as FabricDropdown,
  IDropdownOption,
} from "office-ui-fabric-react";
// import {Dialog} from "@microsoft/sp-dialog";
import * as GlobalConstants from "../../../helperFiles/constants";

export interface IChartWpStates {
  finalArrayCount: any[];
  valueArrayProd: string[];
  valueArrayCount: number[];
  finalArray: any[];
  totalMailCount: number;

  selected_chartType: ChartType;
  defaultSelected_ChartType: number;
}

export default class ChartWp extends React.Component<
  IGoogleChartWpProps,
  IChartWpStates,
  {}
> {
  private category = [
    { key: 0, text: "Bar" },
    { key: 1, text: "Bubble" },
    { key: 2, text: "Doughnut" },
    { key: 3, text: "HorizontalBar" },
    { key: 4, text: "Line" },
    { key: 5, text: "Pie" },
    { key: 6, text: "PolarArea" },
    { key: 7, text: "Radar" },
    { key: 8, text: "Scatter" },
  ];

  private finalset = [];

  constructor(props) {
    super(props);
    this.state = {
      finalArrayCount: [],
      valueArrayProd: [],
      valueArrayCount: [],
      selected_chartType: "doughnut",
      defaultSelected_ChartType: 0,
      finalArray: [],
      totalMailCount: 0
    };

       
  }

  public render(): React.ReactElement<IGoogleChartWpProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
    } = this.props;

    return (
      <section
        className={`${styles.googleChartWp} ${hasTeamsContext ? styles.teams : ""}`}
      >
        <div className={styles.welcome}>
          <div>
            <div>
              {/* <FabricDropdown
              className={styles.DropdownCss}
              placeholder="Chart Type"
              options={this.category}
              onChange={this.categoryChanged}
              defaultSelectedKey={this.state.defaultSelected_ChartType}
              ariaLabel="Please choose chart type"
            /> */}
            </div>

            <table>

              <tr>
                <td width={"100%"} 
                    // className={styles.canvas}
                    >
                  <ChartControl
                    type={"bar"}
                    // data={{
                    //   // labels: this.state.valueArrayProd,
                    //   labels: ['SharePoint', 'OneDrive'],
                    //   datasets: [{
                    //     label: 'My First dataset',
                    //     // data: this.state.valueArrayCount
                    //     data: [20, 15]
                    //   }]
                    // }}
                    className={styles.topSpace}
                    datapromise={this._loadAsyncData()}
                    options={{
                      maintainAspectRatio: false,
                      scales: {
                        yAxes: [
                          {
                            ticks: {
                              beginAtZero: true,
                            },
                          },
                        ],
                      },
                    }}
                    loadingtemplate={() => (
                      <Spinner
                        size={SpinnerSize.large}
                        label="Loading..."
                      ></Spinner>
                    )}
                    rejectedtemplate={(error: string) => (
                      <div>Something went wrong: {error}</div>
                    )}
                    palette={ChartPalette.OfficeColorful3}
                    accessibility={{
                      alternateText:
                        "Text alternative for this canvas graphic is in the data table below.",
                      summary:
                        "This is the text alternative for the canvas graphic.",
                      caption: "Votes for favorite pets",
                    }}
                  />
                </td>
                {/* <td> */}
                  {/* <ChartControl
                    type={"pie"}
                    className={styles.topSpace}
                    // data={{
                    //   // labels: this.state.valueArrayProd,
                    //   labels: ['SharePoint', 'OneDrive'],
                    //   datasets: [{
                    //     label: 'My First dataset',
                    //     // data: this.state.valueArrayCount
                    //     data: [20, 15]
                    //   }]
                    // }}
                    datapromise={this._loadAsyncData()}
                    options={{
                      scales: {
                        yAxes: [
                          {
                            ticks: {
                              beginAtZero: true,
                            },
                          },
                        ],
                      },
                    }}
                    loadingtemplate={() => (
                      <Spinner
                        size={SpinnerSize.large}
                        label="Loading..."
                      ></Spinner>
                    )}
                    rejectedtemplate={(error: string) => (
                      <div>Something went wrong: {error}</div>
                    )}
                    palette={ChartPalette.OfficeColorful1}
                    accessibility={{
                      alternateText:
                        "Text alternative for this canvas graphic is in the data table below.",
                      summary:
                        "This is the text alternative for the canvas graphic.",
                      caption: "Votes for favorite pets",
                    }}
                  />*/}
                {/* </td>  */}
              </tr>

              <tr>
                {/* <td>
                <ChartControl
                    type={"horizontalBar"}
                    className={styles.topSpace}
                    // data={{
                    //   // labels: this.state.valueArrayProd,
                    //   labels: ['SharePoint', 'OneDrive'],
                    //   datasets: [{
                    //     label: 'My First dataset',
                    //     // data: this.state.valueArrayCount
                    //     data: [20, 15]
                    //   }]
                    // }}

                    datapromise={this._loadAsyncData()}
                    options={{
                      scales: {
                        yAxes: [
                          {
                            ticks: {
                              beginAtZero: true,
                            },
                          },
                        ],
                      },
                    }}
                    loadingtemplate={() => (
                      <Spinner
                        size={SpinnerSize.large}
                        label="Loading..."
                      ></Spinner>
                    )}
                    rejectedtemplate={(error: string) => (
                      <div>Something went wrong: {error}</div>
                    )}
                    palette={ChartPalette.OfficeColorful1}
                    accessibility={{
                      alternateText:
                        "Text alternative for this canvas graphic is in the data table below.",
                      summary:
                        "This is the text alternative for the canvas graphic.",
                      caption: "Votes for favorite pets",
                    }}
                  />
                 
                </td>
                <td>
                  
                  <ChartControl
                    type={"doughnut"}
                    className={styles.topSpace}
                    // data={{
                    //   // labels: this.state.valueArrayProd,
                    //   labels: ['SharePoint', 'OneDrive'],
                    //   datasets: [{
                    //     label: 'My First dataset',
                    //     // data: this.state.valueArrayCount
                    //     data: [20, 15]
                    //   }]
                    // }}

                    datapromise={this._loadAsyncData()}
                    options={{
                      scales: {
                        yAxes: [
                          {
                            ticks: {
                              beginAtZero: true,
                            },
                          },
                        ],
                      },
                    }}
                    loadingtemplate={() => (
                      <Spinner
                        size={SpinnerSize.large}
                        label="Loading..."
                      ></Spinner>
                    )}
                    rejectedtemplate={(error: string) => (
                      <div>Something went wrong: {error}</div>
                    )}
                    palette={ChartPalette.OfficeColorful1}
                    accessibility={{
                      alternateText:
                        "Text alternative for this canvas graphic is in the data table below.",
                      summary:
                        "This is the text alternative for the canvas graphic.",
                      caption: "Votes for favorite pets",
                    }}
                  />
                </td> */}
              </tr>

              <tr>
                {/* <td colSpan={2}>
                  {" "}
                  <ChartControl
                    type={"line"}
                    className={styles.topSpace}
                    // data={{
                    //   // labels: this.state.valueArrayProd,
                    //   labels: ['SharePoint', 'OneDrive'],
                    //   datasets: [{
                    //     label: 'My First dataset',
                    //     // data: this.state.valueArrayCount
                    //     data: [20, 15]
                    //   }]
                    // }}

                    datapromise={this._loadAsyncData()}
                    options={{
                      scales: {
                        yAxes: [
                          {
                            ticks: {
                              beginAtZero: true,
                            },
                          },
                        ],
                      },
                    }}
                    loadingtemplate={() => (
                      <Spinner
                        size={SpinnerSize.large}
                        label="Loading..."
                      ></Spinner>
                    )}
                    rejectedtemplate={(error: string) => (
                      <div>Something went wrong: {error}</div>
                    )}
                    palette={ChartPalette.OfficeColorful1}
                    accessibility={{
                      alternateText:
                        "Text alternative for this canvas graphic is in the data table below.",
                      summary:
                        "This is the text alternative for the canvas graphic.",
                      caption: "Votes for favorite pets",
                    }}
                  />
                </td> */}
              </tr>

            </table>
          </div>

          <div>
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
          </ul>*/}
            <div id='table' >
                    
                    {/* <table>
                      <tr><td colSpan={2}><span style={{fontWeight:"bold"}}>Total :</span> {this.state.totalMailCount}</td></tr>
                     
                    </table>
                    <br></br>
                    <table>
                      <tr>
                        
                        {this.finalset.map(column => <><td ><span style={{fontWeight:"bold"}}>{column.Tag} :</span> {column.occurrence}</td></>)}
                      </tr>
                    </table> */}
                  
            </div>
          </div>
        </div>
      </section>
    );
  }

  public async componentDidMount() {
    // this._loadAsyncData_1();
    // this.getDataFromNBHCategoryList();
    // this.getDatafromSharePointList();   // To check users exist in which all groups
    // //get current logged in user details
    // await sp.web.currentUser.get().then((r) => { this.email = r.Email; this.displayName = r.Title; });
    // this.setState({ bookedFor: this.displayName });

    // document.getElementById("spLeftNav").style.display ="none";

    // Dialog.prompt("abc");
  }

  //onchange event of rooms dropdown - Add Recurring
  private categoryChanged = (
    event: React.FormEvent<HTMLDivElement>,
    item: IDropdownOption
  ): void => {
    this.setState({ defaultSelected_ChartType: Number(item.id) });

    if (item.text == "Bar") {
      this.setState({ selected_chartType: "bar" });
    } else if (item.text == "Bubble") {
      this.setState({ selected_chartType: "bubble" });
    } else if (item.text == "Doughnut") {
      this.setState({ selected_chartType: "doughnut" });
    } else if (item.text == "HorizontalBar") {
      this.setState({ selected_chartType: "horizontalBar" });
    } else if (item.text == "Line") {
      this.setState({ selected_chartType: "line" });
    } else if (item.text == "Pie") {
      this.setState({ selected_chartType: "pie" });
    } else if (item.text == "PolarArea") {
      this.setState({ selected_chartType: "polarArea" });
    } else if (item.text == "Radar") {
      this.setState({ selected_chartType: "radar" });
    } else if (item.text == "Scatter") {
      this.setState({ selected_chartType: "scatter" });
    }

    // this.setState({ (item.text == "Bar") })
    // this.setState({ selected_chartType: item.text });
    console.log(
      "this.state.selected_chartType - ",
      this.state.selected_chartType
    );

    // return this.state.selected_chartType
  }

  private _loadAsyncData = async () => {
    // private getDatafromSharePointList = async () => {
    // Connection to the current context's Web
    // const sp = spfi(this.context);

    // Get all items from List
    const res_AllListData_Array = await sp.web.lists
      .getByTitle(GlobalConstants.lstName_productSupport)
      .items.select("*")
      .getAll();
    console.log("Result : ", res_AllListData_Array);

    //Push all data into required array object
    let AllListData_Array: any[] = [];

    let AllListData_Array_filter = res_AllListData_Array.filter((data) => data.Tags == "Google");
    console.log("AllListData_Array_filter : ", AllListData_Array_filter);


    AllListData_Array_filter.forEach((element) => {
      AllListData_Array.push({
        ID: element.ID,
        text: element.Title,
        SubTags: element.SubTags,
      });
    });
    console.log("valueArray : ", AllListData_Array);

    //find duplicate items count
    let finalset = await this.findOcc(AllListData_Array, "SubTags");
    // this.setState({finalArrayCount : finalset});
    console.log("finalset - ", finalset);

    // this.setState({finalArrayCount : finalset});

    let ProductsArray_lbl: string[] = [];
    let ProductsCountArray_value: number[] = [];

    finalset.forEach((element) => {
      ProductsArray_lbl.push(element.SubTags);
      ProductsCountArray_value.push(element.occurrence);
    });

    let chartdata: any = {
      labels: ProductsArray_lbl,
      datasets: [
        {
          label: "Google mails Report",
          data: ProductsCountArray_value,
        },
      ],
    };
    return chartdata;
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
    // let 
    this.finalset = await this.findOcc(AllListData_Array, "Tag") ;
    // this.setState({finalArrayCount : finalset});
    console.log("finalset - ", this.finalset);
  
  
    // this.setState({finalArray : finalset, totalMailCount: res_AllListData_Array.length});
  

          
  
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

  // private _loadAsyncData_test(): Promise<Chart.ChartData> {
  //   return new Promise<Chart.ChartData>((resolve, reject) => {
  //     // Call your own service -- this example returns an array of numbers
  //     // but you could call
  //     const dataProvider: IChartDataProvider = new MockChartDataProvider();
  //     dataProvider.getNumberArray().then((numbers: number[]) => {
  //       // format your response to ChartData
  //       const data: Chart.ChartData =
  //       {
  //         labels: ['January', 'February', 'March', 'April', 'May', 'June', 'July']
  //         datasets: [
  //           {
  //             label: 'My First dataset',
  //             data: numbers
  //           }
  //         ]
  //       };

  //       // resolve the promise
  //       resolve(data);
  //     });
  //   });
  // }

  // private async _loadAsyncData(): Promise<any> {
  //   const items: any[] = await sp.web.lists.getByTitle("Sales").items.select("Title", "Sales").get();
  //   let lblarr: string[] = [];
  //   let dataarr: number[] = [];

  //     items.forEach(element => {
  //       lblarr.push(element.Title);
  //       dataarr.push(element.Sales);
  //     });

  //   let chartdata: any = {
  //     labels: lblarr,
  //     datasets: [{
  //       label: 'My Sales',
  //       data: dataarr
  //     }]
  //   };
  //   return chartdata;
  // }

  //get data from NBH category list
}
