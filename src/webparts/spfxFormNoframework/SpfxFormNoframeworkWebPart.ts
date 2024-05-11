import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
//import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SpfxFormNoframeworkWebPart.module.scss';
import * as strings from 'SpfxFormNoframeworkWebPartStrings';
import {ISPHttpClientOptions, SPHttpClient,SPHttpClientResponse} from '@microsoft/sp-http';


export interface ISpfxWebpartNoFrameworkWebPartProps {
  description: string;
}

export default class SpfxWebpartNoFrameworkWebPart extends BaseClientSideWebPart<ISpfxWebpartNoFrameworkWebPartProps> {

//  private _isDarkTheme: boolean = false;
//  private _environmentMessage: string = '';

  public render(): void {
    this.domElement.innerHTML = `
    <div>

        
    
        <!-- Basic details -->
        <div class=${styles.welcome} id="BasicDetails">

          <p> Please the fill basic details of the contract </p>

          <input type="text" id="contractNumber" placeholder="Contract Number" name="Contract Number" />
          <input type="text" id="nextAction" placeholder="Next Action" name="Next Action" />
          <input type="text" id="customerName" placeholder="Customer Name" name="Customer Name" />
          <input type="text" id="industrySegment" placeholder="Industry Segment" name="Industry Segment" />
          <input type="text" id="salesUnit" placeholder="Sales Unit" name="Sales Unit" />
          <input type="text" id="salesPerson" placeholder="Sales Person" name="Sales Person" />
          <input type="text" id="reason" placeholder="Reason" name="Reason" />
          <input type="text" id="product" placeholder="Product" name="Product" />
          <input type="text" id="lastChange" placeholder="Last Change" name="Last Change" />
          <input type="text" id="comments" placeholder="Comments" name="Comments" />
        
        </div>
        
        <!-- Template Selection -->
        <div id="TemplateSelection">

          <label for="country">Choose a country Template</label>

          <select type="text" id="country">
            <option value="UK">United Kingdom</option>
            <option value="ES">Spain</option>
            <option value="US">United States</option>
          </select>
        
        </div>
        
        <!-- Approver -->
        <div id="Approver">
          <p> Please the fill the approver </p>
          <input type="text" id="approver" placeholder="Approver" name="Approver" />

        </div>
        
        
    
      <br/>

        <div>
            <!-- Submit Button -->
           
                  
            <!-- PaginationButtons -->
            <div id ="NavigationButtonsLeft">
            <input  type="button" id="BackBasicDetails" value="Back - Basic Details"></input>  
            <input type="button" id="BackTemplate" value="Back - Choose Template"></input> 

            </div>
            <div id ="NavigationButtonsRight">
            <input type="button"  id="NextBasicDetails" value="Next - Template Selection"></input>
            <input type="button" id="NextTemplate" value="Next - Approvers"></input> 
            <input type="button" id="BttnEmp" value="Submit"></input> 
            </div>
        </div>
      
      
      </div>`;

      this._bindSave();
      this._initialLayout();
      this._paginationNextBasicDetails();
      this._paginationNextTemplate();
      this._backBasicDetails();
      this._backTemplate();
  }

  private _initialLayout(): void {
    // Pagination Next Basic Details (Basic details --> Template Selection) & BreadCrumb Update 
    const NextBasicDetails = this.domElement.querySelector('#NextBasicDetails');
       if (NextBasicDetails) {

        // Elements to hide

            // Hide BttnEmp
            var BttnEmp = document.getElementById("BttnEmp");
            if (BttnEmp !== null) {
              BttnEmp.style.display = "none";
            } else {
                // Handle the case where the element is not found
                console.error("Element with id 'BttnEmp' not found");
            }

            // Hide NextTemplate
             var NextTemplate = document.getElementById("NextTemplate");
             if (NextTemplate !== null) {
              NextTemplate.style.display = "none";
             } else {
                 // Handle the case where the element is not found
                 console.error("Element with id 'NextTemplate' not found");
             }

            // Hide BackBasicDetails
            var BackBasicDetails = document.getElementById("BackBasicDetails");
            if (BackBasicDetails !== null) {
              BackBasicDetails.style.display = "none";
            } else {
                // Handle the case where the element is not found
                console.error("Element with id 'BackBasicDetails' not found");
            }

            // Hide BackTemplate
            var BackTemplate = document.getElementById("BackTemplate");
            if (BackTemplate !== null) {
              BackTemplate.style.display = "none";
            } else {
                // Handle the case where the element is not found
                console.error("Element with id 'BackTemplate' not found");
            }

            // Hide Approver
            var Approver = document.getElementById("Approver");
            if (Approver !== null) {
              Approver.style.display = "none";
            } else {
                // Handle the case where the element is not found
                console.error("Element with id 'Approver' not found");
            }

            // Hide TemplateSelection
            var TemplateSelection = document.getElementById("TemplateSelection");
            if (TemplateSelection !== null) {
              TemplateSelection.style.display = "none";
            } else {
                // Handle the case where the element is not found
                console.error("Element with id 'TemplateSelection' not found");
            }
        
       } else {
           console.error("Button element '#BttnEmp' not found.");
         }

 }



  private _paginationNextBasicDetails(): void {
     // Pagination Next Basic Details (Basic details --> Template Selection) & BreadCrumb Update 
     const NextBasicDetails = this.domElement.querySelector('#NextBasicDetails');
        if (NextBasicDetails) {

          NextBasicDetails.addEventListener('click', () => { 

            // Elements to hide

            // Hide NextBasicDetails
            var NextBasicDetail = document.getElementById("NextBasicDetails");
            if (NextBasicDetail !== null) {
              NextBasicDetail.style.display = "none";
            } else {
                // Handle the case where the element is not found
                console.error("Element with id 'NextBasicDetails' not found");
            }

            // Hide BasicDetails
            var BasicDetails = document.getElementById("BasicDetails");
            if (BasicDetails !== null) {
              BasicDetails.style.display = "none";
            } else {
                // Handle the case where the element is not found
                console.error("Element with id 'BasicDetails' not found");
            }

            // Hide Approver
            var Approver = document.getElementById("Approver");
            if (Approver !== null) {
              Approver.style.display = "none";
            } else {
                // Handle the case where the element is not found
                console.error("Element with id 'Approver' not found");
            }

            // Hide BttnEmp
            var BttnEmp = document.getElementById("BttnEmp");
            if (BttnEmp !== null) {
              BttnEmp.style.display = "none";
            } else {
                // Handle the case where the element is not found
                console.error("Element with id 'BttnEmp' not found");
            }

            // Elements to show
            
            // Show TemplateSelection
            var TemplateSelection = document.getElementById("TemplateSelection");
            if (TemplateSelection !== null) {
              TemplateSelection.style.display = "Block";
            } else {
                // Handle the case where the element is not found
                console.error("Element with id 'TemplateSelection' not found");
            }

            // Show NextTemplate
            var NextTemplate = document.getElementById("NextTemplate");
            if (NextTemplate !== null) {
              NextTemplate.style.display = "Block";
            } else {
                // Handle the case where the element is not found
                console.error("Element with id 'NextTemplate' not found");
            }

            // Show BackBasicDetails
            var BackBasicDetails = document.getElementById("BackBasicDetails");
            if (BackBasicDetails !== null) {
              BackBasicDetails.style.display = "Block";
            } else {
                // Handle the case where the element is not found
                console.error("Element with id 'BackBasicDetails' not found");
            }
            
        });
          
        } else {
            console.error("Button element '#NextBasicDetails' not found.");
          }

  }


  private _paginationNextTemplate(): void {
     // Pagination from NextTemplate (Template Selection to Approvals) & BreadCrumb Update 
    const NextTemplate = this.domElement.querySelector('#NextTemplate');
       if (NextTemplate) {

        NextTemplate.addEventListener('click', () => { 

           // Elements to hide

           // Hide NextBasicDetails
           var NextBasicDetail = document.getElementById("NextBasicDetails");
           if (NextBasicDetail !== null) {
             NextBasicDetail.style.display = "none";
           } else {
               // Handle the case where the element is not found
               console.error("Element with id 'NextBasicDetails' not found");
           }

           // Hide BasicDetails
           var BasicDetails = document.getElementById("BasicDetails");
           if (BasicDetails !== null) {
             BasicDetails.style.display = "none";
           } else {
               // Handle the case where the element is not found
               console.error("Element with id 'BasicDetails' not found");
           }

           // Hide TemplateSelection
           var TemplateSelection = document.getElementById("TemplateSelection");
           if (TemplateSelection !== null) {
            TemplateSelection.style.display = "none";
           } else {
               // Handle the case where the element is not found
               console.error("Element with id 'TemplateSelection' not found");
           }

           // Hide NextTemplate
           var NextTemplate = document.getElementById("NextTemplate");
           if (NextTemplate !== null) {
            NextTemplate.style.display = "none";
           } else {
               // Handle the case where the element is not found
               console.error("Element with id 'NextTemplate' not found");
           }

          // Hide BackBasicDetails
          var BackBasicDetails = document.getElementById("BackBasicDetails");
          if (BackBasicDetails !== null) {
            BackBasicDetails.style.display = "none";
          } else {
              // Handle the case where the element is not found
              console.error("Element with id 'BackBasicDetails' not found");
          }


           // Elements to show
           
           // Show Approver
           var Approver = document.getElementById("Approver");
           if (Approver !== null) {
            Approver.style.display = "Block";
           } else {
               // Handle the case where the element is not found
               console.error("Element with id 'Approver' not found");
           }

           // Show BttnEmp
           var BttnEmp = document.getElementById("BttnEmp");
           if (BttnEmp !== null) {
            BttnEmp.style.display = "Block";
           } else {
               // Handle the case where the element is not found
               console.error("Element with id 'BttnEmp' not found");
           }

           // Show BackTemplate
           var BackTemplate = document.getElementById("BackTemplate");
           if (BackTemplate !== null) {
            BackTemplate.style.display = "Block";
           } else {
               // Handle the case where the element is not found
               console.error("Element with id 'BackTemplate' not found");
           }
           
       });
         
       } else {
           console.error("Button element '#NextTemplate' not found.");
         }

  }


  private _backBasicDetails(): void {
      // Pagination Back Basic Details (Template Selection --> Basic details ) & BreadCrumb Update 
      const BackBasicDetails = this.domElement.querySelector('#BackBasicDetails');
      if (BackBasicDetails) {

        BackBasicDetails.addEventListener('click', () => { 

          // Elements to hide

          // Hide NextBasicDetails
          var NextBasicDetail = document.getElementById("NextBasicDetails");
          if (NextBasicDetail !== null) {
            NextBasicDetail.style.display = "none";
          } else {
              // Handle the case where the element is not found
              console.error("Element with id 'NextBasicDetails' not found");
          }

          
          // Hide TemplateSelection
          var TemplateSelection = document.getElementById("TemplateSelection");
          if (TemplateSelection !== null) {
          TemplateSelection.style.display = "none";
          } else {
              // Handle the case where the element is not found
              console.error("Element with id 'TemplateSelection' not found");
          }

          // Hide Approver
          var Approver = document.getElementById("Approver");
          if (Approver !== null) {
            Approver.style.display = "none";
          } else {
              // Handle the case where the element is not found
              console.error("Element with id 'Approver' not found");
          }

          
          
          // Hide BttnEmp
          var BttnEmp = document.getElementById("BttnEmp");
          if (BttnEmp !== null) {
            BttnEmp.style.display = "none";
          } else {
              // Handle the case where the element is not found
              console.error("Element with id 'BttnEmp' not found");
          }
          
          // Hide BackBasicDetails
          var BackBasicDetails = document.getElementById("BackBasicDetails");
          if (BackBasicDetails !== null) {
            BackBasicDetails.style.display = "none";
          } else {
              // Handle the case where the element is not found
              console.error("Element with id 'BackBasicDetails' not found");
          }

          // Hide NextTemplate
          var NextTemplate = document.getElementById("NextTemplate");
          if (NextTemplate !== null) {
          NextTemplate.style.display = "none";
          } else {
              // Handle the case where the element is not found
              console.error("Element with id 'NextTemplate' not found");
          }


          // Elements to show
          
          // Show BasicDetails
          var BasicDetails = document.getElementById("BasicDetails");
          if (BasicDetails !== null) {
          BasicDetails.style.display = "Block";
          } else {
              // Handle the case where the element is not found
              console.error("Element with id 'BasicDetails' not found");
          }

          // Show NextBasicDetails
          var NextBasicDetails = document.getElementById("NextBasicDetails");
          if (NextBasicDetails !== null) {
          NextBasicDetails.style.display = "Block";
          } else {
              // Handle the case where the element is not found
              console.error("Element with id 'NextBasicDetails' not found");
          }

          
          
      });
        
      } else {
          console.error("Button element '#BackBasicDetails' not found.");
        }

  }


  private _backTemplate(): void {
    // Pagination from NextTemplate (Template Selection to Approvals) & BreadCrumb Update 
    const BackTemplate = this.domElement.querySelector('#BackTemplate');
    if (BackTemplate) {

      BackTemplate.addEventListener('click', () => { 

        // Elements to hide

        // Hide NextBasicDetails
        var NextBasicDetail = document.getElementById("NextBasicDetails");
        if (NextBasicDetail !== null) {
          NextBasicDetail.style.display = "none";
        } else {
            // Handle the case where the element is not found
            console.error("Element with id 'NextBasicDetails' not found");
        }

        
        // Hide BasicDetails
        var BasicDetails = document.getElementById("BasicDetails");
        if (BasicDetails !== null) {
          BasicDetails.style.display = "none";
        } else {
            // Handle the case where the element is not found
            console.error("Element with id 'BasicDetails' not found");
        }

        // Hide Approver
        var Approver = document.getElementById("Approver");
        if (Approver !== null) {
          Approver.style.display = "none";
        } else {
            // Handle the case where the element is not found
            console.error("Element with id 'Approver' not found");
        }

        
        
        // Hide BttnEmp
        var BttnEmp = document.getElementById("BttnEmp");
        if (BttnEmp !== null) {
          BttnEmp.style.display = "none";
        } else {
            // Handle the case where the element is not found
            console.error("Element with id 'BttnEmp' not found");
        }
        
        // Hide BackTemplate
        var BackTemplate = document.getElementById("BackTemplate");
        if (BackTemplate !== null) {
          BackTemplate.style.display = "none";
        } else {
            // Handle the case where the element is not found
            console.error("Element with id 'BackTemplate' not found");
        }

        


        // Elements to show
        
        // Show TemplateSelection
        var TemplateSelection = document.getElementById("TemplateSelection");
        if (TemplateSelection !== null) {
          TemplateSelection.style.display = "Block";
        } else {
            // Handle the case where the element is not found
            console.error("Element with id 'BasicDetails' not found");
        }

        // Show NextTemplate
        var NextTemplate = document.getElementById("NextTemplate");
        if (NextTemplate !== null) {
          NextTemplate.style.display = "Block";
        } else {
            // Handle the case where the element is not found
            console.error("Element with id 'NextTemplate' not found");
        }

        // Show BackBasicDetails
        var BackBasicDetails = document.getElementById("BackBasicDetails");
        if (BackBasicDetails !== null) {
          BackBasicDetails.style.display = "Block";
        } else {
            // Handle the case where the element is not found
            console.error("Element with id 'BackBasicDetails' not found");
        }
        
        
    });
      
    } else {
        console.error("Button element '#BackTemplate' not found.");
      }

  }


  private _bindSave(): void {

    const button = this.domElement.querySelector('#BttnEmp');
    if (button) {
        button.addEventListener('click', () => { this.addListItem(); });
    } else {
        console.error("Button element '#BttnEmp' not found.");
    }
  
  }

  private addListItem(): void{
    
    
    var contractNumber = (document.getElementById("contractNumber") as HTMLInputElement).value;
    var country = (document.getElementById("country") as HTMLInputElement).value;
    //var nextAction = (document.getElementById("nextAction") as HTMLInputElement).value;
    //var customerName = (document.getElementById("customerName") as HTMLInputElement).value;
    //var industrySegment = (document.getElementById("industrySegment") as HTMLInputElement).value;
    //var salesUnit = (document.getElementById("salesUnit") as HTMLInputElement).value;
    //var salesSegment = (document.getElementById("salesSegment") as HTMLInputElement).value;
    //var approver = (document.getElementById("approver") as HTMLInputElement).value;
    //var reason = (document.getElementById("reason") as HTMLInputElement).value;
    //var product = (document.getElementById("product") as HTMLInputElement).value;
    //var lastChange = (document.getElementById("lastChange") as HTMLInputElement).value;
    //var comments = (document.getElementById("comments") as HTMLInputElement).value;
    console.log(contractNumber);
    console.log(country);
    const siteUrl: string = "https://t8656.sharepoint.com/sites/Sharepoint_Interaction/_api/web/lists/getbytitle('PoC_ContractHUB2')/items"
    const itemBody: any = {

      "Title": 'TestWebpart',
      "Description": 'TestWebPart'
    };

    const spHttpClientOptions: ISPHttpClientOptions = {
      "body": JSON.stringify(itemBody)
    };

    this.context.spHttpClient.post(siteUrl,SPHttpClient.configurations.v1,spHttpClientOptions)
    .then((response: SPHttpClientResponse)=> {
      alert("Success!!");
    });
}
  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
   //   this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

   // this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
