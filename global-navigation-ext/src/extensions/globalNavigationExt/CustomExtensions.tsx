import * as React from 'react';

import { IntranetLocation, IntranetTrigger, IntranetProvider, ExtensionService, TriggerService, ProviderService, ExtensionProvider, IUserProfileProvider, DataSourceService, ExtensionPointToolboxAction, ExtensionPointToolboxPanelCreationAction, MegaMenuItem, StorageType, IClientStorageProvider, IMyToolsProvider, INavigationHierarchyProvider, MegaMenuNavigationItem, InformationMessage, ContextActionType } from '@valo/extensibility';
import { IMultilingualProvider } from '@valo/extensibility/lib/providerTypes/IMultilingualProvider';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { sp } from "@pnp/sp";
import { Web, IWeb } from "@pnp/sp/webs";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/fields/list";
import { Camera, Plus, ArrowRight } from 'react-feather';
import './notify.css';
// import Clock from './clock';
// import { NoPagingDataSource } from './datasource/NoPagingDataSource';
// import { DynamicPagingDataSource } from './datasource/DynamicPagingDataSource';
// import { StaticDataSource } from './datasource/StaticDataSource';
import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';
// import { PersonalNavigation } from './navigation/PersonalNavigation';
// import { StaticNavigation } from './navigation/StaticNavigation';
// import { CustomNavigationItem } from './customNavigationItem';
// import { CustomGroupHeader } from './customGroupHeader';
// import { Footer } from './footer';
// import { CustomNotifications } from './customNotifications/CustomNotifications';
// import ToolboxComponent from './toolboxComponent/ToolboxComponent';

import ColorPickerDialog from './ColorPickerDialog';
import { IColor } from 'office-ui-fabric-react';
import ShowItemDetails from './ShowItemDetails';

import { NotificationDetails } from './notificationDetails'
import AddItemDetails from './AddItemPopup';

export const CustomBreadcrumb: React.SFC<any> = (props: any) => {
  console.log('CustomBreadcrumb', props);
  return (
    <div>
      {
        Object.keys(props).map(key => <span>{key}: {props[key]}</span>)
      }
    </div>
  );
};

export default class CustomExtensions {
  private extensionService: ExtensionService = null;
  private triggerService: TriggerService = null;
  private providerService: ProviderService = null;
  private dataSourceService: DataSourceService = null;

  constructor() {
    this.extensionService = ExtensionService.getInstance();
    this.triggerService = TriggerService.getInstance();
    this.providerService = ProviderService.getInstance();
    this.dataSourceService = DataSourceService.getInstance();
  }

  public register(ctx: ApplicationCustomizerContext) {
    this.ListProvisionOrDataPopulate(ctx);
  }


  private async ListProvisionOrDataPopulate(ctx: ApplicationCustomizerContext) {
    var adminUrl = ctx.pageContext.web.absoluteUrl.replace(ctx.pageContext.web.serverRelativeUrl, "");
    adminUrl = adminUrl + "/sites/valoadmin";

    // Provision SP list if it's not available already. Will be executed while adding the extension to menu bar.
    // User needs to have edit rights in the site for list provision.
    var adminWeb = Web(adminUrl);
    adminWeb.lists.getByTitle("NotificationConfig").
      items.select("Title").top(1).
      orderBy("Modified", true).
      get().
      then((items: any) => {
        if (items.length > 0) {
          var siteUrl = items[0].Title;
          var siteWeb = Web(siteUrl);

          /// Change by Arijit :: start
          const d = new Date();
          console.log("Adjustment for Visitor !!!")
          // If list already exists in the site get data to display from notification
          siteWeb.lists.getByTitle("NotificationList").
            items.select("Title, Description, StartDate, EndDate, SalesNo").
            filter(`EndDate ge datetime'${d.toISOString()}'`).
            get().
            then((items: any) => {
              this.RegisterNotificationIconAndPopulateData(items, ctx);
              console.log(items);
            });

            // Change by Arijit :: End



          /* const lst = siteWeb.lists.ensure("NotificationList");

          lst.then((listExistResult) => {
            if (listExistResult.created) {
              console.log("List created!");
              const listNotif = siteWeb.lists.getByTitle("NotificationList");
              console.log("Description rich text !!");
              listNotif.fields.createFieldAsXml(
                '<Field DisplayName="Description"  Type="Note" Required="FALSE" StaticName="Description" Name="Description"/>'
              //  '<Field Type="Note" DisplayName="Description" Required="FALSE"  NumLines="6" RichText="TRUE" RichTextMode="FullHtml" IsolateStyles="TRUE" Sortable="FALSE" AppendOnly="FALSE" StaticName="Description" Name="Description" Description="Please provide short description"/>'
              ).then((a) => {
                listNotif.fields.createFieldAsXml(
                  '<Field DisplayName="StartDate" Type="DateTime" Required="FALSE" StaticName="StartDate" Name="StartDate"/>'
                ).then((b) => {
                  listNotif.fields.createFieldAsXml(
                    '<Field DisplayName="EndDate" Type="DateTime" Required="FALSE" StaticName="EndDate" Name="EndDate"/>'
                  ).then((c) => {
                    listNotif.fields.createFieldAsXml(
                      '<Field DisplayName="SalesNo" Type="Text" Required="FALSE" StaticName="SalesNo" Name="SalesNo"/>'
                    )
                  })
                })
              });
            }
            else {
              console.log("List already existed!");
              const d = new Date();

              // If list already exists in the site get data to display from notification
              siteWeb.lists.getByTitle("NotificationList").
                items.select("Title, Description, StartDate, EndDate, SalesNo").
                filter(`EndDate ge datetime'${d.toISOString()}'`).
                get().
                then((items: any) => {
                  this.RegisterNotificationIconAndPopulateData(items, ctx);
                  console.log(items);
                });
            }
          });*/
        }
      });

    // End list provision checking and data population
  }

  private RegisterNotificationIconAndPopulateData(items: any, ctx: ApplicationCustomizerContext) {
    this.extensionService.registerExtension({
      id: "NavigationRight",
      location: IntranetLocation.NavigationRight,
      element: <div id="divNotification"  className="arrow_box" style={{ cursor: 'pointer', padding: '17px 9px 15px 8px' }}>
        <div style={{ textAlign: 'center', position:'relative' }}>
          <svg
            width="24"
            height="24"
            viewBox="0 0 24 24"
            fill="red"
            xmlns="http://www.w3.org/2000/svg"
            style={{ display: 'inline-block' }}
            id="svgImg"
            
          >
            <path
              fill-rule="evenodd"
              clip-rule="evenodd"
              d="M14 3V3.28988C16.8915 4.15043 19 6.82898 19 10V17H20V19H4V17H5V10C5 6.82898 7.10851 4.15043 10 3.28988V3C10 1.89543 10.8954 1 12 1C13.1046 1 14 1.89543 14 3ZM7 17H17V10C17 7.23858 14.7614 5 12 5C9.23858 5 7 7.23858 7 10V17ZM14 21V20H10V21C10 22.1046 10.8954 23 12 23C13.1046 23 14 22.1046 14 21Z"
              fill="currentColor"
            />
          </svg>
          <span style={{
            background: 'blue',
            borderRadius: '10px',
            padding: '3px',
            fontSize: '10px',
            position: 'absolute',
            top: '-5px',
            right: '0'
          }} id="spnCount">
            {items.length}</span></div>
        <div id="divNotificationDrp"
          style={{
            position: 'absolute',
            top: '100%',
            background: 'rgb(255, 255, 255)',
            color: 'rgb(51, 51, 51)',
           //color: 'rgb(244, 242, 230)',
            width: '100%',
            left: '0px',
            boxShadow: 'rgb(0 0 0 / 15%) 0px 3px 5px',
            display: 'none'
          }}>
          <div style={{  background:'#F4F2E6', color: '#000000', padding: '12px 8px', minHeight: '20px' }}>
            Driftsmeldinger
               <button id="btnShow"  className="showBtn addButtonPopup"> Show All</button>
                <button id="btnAdd" className="showBtn addButtonPopup" style={{marginRight: '5px'}}><Plus size={14} /> Add</button>
          </div>
          <div id="divNotificationDrpinner" style={{ padding: '2rem', fontSize: '1em' }}>
            <NotificationDetails name="susanta" items={items} ctx={ctx} />
          </div>
        </div>
      </div>
    });

    let clickEvent = document.getElementById('divNotification');
    const Area = document.getElementById('divNotificationDrp');
    const btnShow = document.getElementById('btnShow');
    //clickEvent.addEventListener("click", (e: Event) => this.ViewNotification(liItems));
    btnShow.onclick = function(){
      window.open("https://vtfk.sharepoint.com/sites/innsida/Lists/NotificationList/AllItems.aspx");
    }
    clickEvent.onclick = function (e) {
      console.log('clicked now');
      //clickEvent.classList.add("hovered");
     
      if (Area.style.display === "none") {
        Area.style.display = "block";
      } else {
        Area.style.display = "none";
      }
     
    };
    
     
   /* clickEvent.onmouseleave = function (e) {
      //console.log('mout')
      setTimeout(function () { Area.style.display = "none"; }, 500);
      //Area.style.display = "none";
      //clickEvent.classList.remove("hovered");
    };*/

    let btnAddClick = document.getElementById('btnAdd');
    btnAddClick.addEventListener("click", (e: Event) => this.ShowItemAddPopup(ctx));
  }

  private ShowItemAddPopup(ctx: ApplicationCustomizerContext) {
    const dialog: AddItemDetails = new AddItemDetails();
    dialog.ctx = ctx;
    dialog.show().then(() => {
      //
    });
  }

  private async fetchMultilingualInformation() {
    const multilingualProvider = await this.providerService.getProvider<IMultilingualProvider>(IntranetProvider.Multilingual);
    if (multilingualProvider && multilingualProvider.instance) {
      const crntPage = await multilingualProvider.instance.getCurrentPage();
      if (crntPage && crntPage.UniqueId) {
        const pages = await multilingualProvider.instance.getPageConnections(crntPage.UniqueId);
        const sites = await multilingualProvider.instance.getSiteConnections();
        const languageTerms = await multilingualProvider.instance.getLanguageTerms();

        console.log(pages, sites, languageTerms);
      }
    }
  }

  private async fetchMyTools() {
    const myToolsProvider = await this.providerService.getProvider<IMyToolsProvider>(IntranetProvider.MyTools);

    if (myToolsProvider && myToolsProvider.instance) {
      const myToolsInstance = myToolsProvider.instance;
      console.log(await myToolsInstance.getMyLinks(25, 0));
      console.log(await myToolsInstance.getOurLinks(25, 0));
    }
  }

  private async fetchNavigation() {
    const navProvider = await this.providerService.getProvider<INavigationHierarchyProvider>(IntranetProvider.NavigationHierarchy);

    if (navProvider && navProvider.instance) {
      const navInstance = navProvider.instance;
      // The navigation provider - provides changes via event emitting. In order to get the new changes, you have to provide the `getHierarchy` method a callback function.
      navInstance.getHierarchy("ExtensibilitySample", (hierarchy: MegaMenuNavigationItem[]) => {
        console.log(hierarchy);
      });
    }
  }
}
