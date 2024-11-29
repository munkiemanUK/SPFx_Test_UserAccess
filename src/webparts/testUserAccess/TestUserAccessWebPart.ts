import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './TestUserAccessWebPart.module.scss';
import * as strings from 'TestUserAccessWebPartStrings';

import { spfi, SPFx as spSPFx } from "@pnp/sp";
import { graphfi, GraphFI, SPFx as graphSPFx } from "@pnp/graph";
import { IFileInfo } from "@pnp/sp/files";
import { LogLevel, PnPLogging } from "@pnp/logging";
import { Group } from "@microsoft/microsoft-graph-types";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/graph/Users";
import "@pnp/graph/groups";

import { SPComponentLoader } from '@microsoft/sp-loader';
import { HttpClient, HttpClientResponse } from '@microsoft/sp-http';

require("bootstrap");

export interface ITestUserAccessWebPartProps {
  description: string;
  userEmail: string;
  libraryUrl: string;
}

export default class TestUserAccessWebPart extends BaseClientSideWebPart<ITestUserAccessWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private sp: ReturnType<typeof spfi>;
  private graph: GraphFI;

  public async getGroupsGraph(): Promise<string[]> {
    //const tenant_uri = this.context.pageContext.web.absoluteUrl.split('/',3)[2];
    let groupNames:any=[];

    try {
      const groups = await this.graph.me.memberOf();
      //const userGroups = await sp.web.siteUsers.getByEmail(userEmail).groups();
      //const userGroups = await sp.web.currentUser.groups();
      //const groupNames = userGroups.map((group: { Title: any; }) => group.Title);
      groupNames = groups
        .filter(group => (group as Group).displayName !== undefined)
        .map(group => (group as Group).displayName);
      
      return groupNames;      

    } catch (error) {
      console.error("Error fetching user groups:", error);
      throw error;
    }
  }

  private getGroupsSP(): Promise<any[]> {
    
    return this.context.httpClient
      .get(`${this.context.pageContext.web.absoluteUrl}/_api/web/currentuser/?$expand=Groups`, HttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json;odata=nometadata', // Request JSON format
        },
      }) 
        .then((response: HttpClientResponse) => {
          if(response.ok){
            console.log("response",response.json());
            return response.json();
          }
      })
      .then((data) => {
        return data.Groups || [];
      });
  }

  public renderGraphGroups(groupNames: string[]): void {
    const groups : Element | null = this.domElement.querySelector("#graphContainer");
    let html:string="";

    groupNames.forEach(groupName => {
      html+=`<li>Graph Title: ${groupName}</li>`;
    });

    if(groups) {
      groups.innerHTML = html;
    }
  }

  public renderSPGroups(groupNames: any): void {
    const groups : Element | null = this.domElement.querySelector("#spContainer");
    let html:string="";

    console.log("groups",groupNames);

    groupNames.forEach((groupName: any) => {
      html+=`<li>Group Title: ${groupName.Title}</li>`;
    });

    if(groups) {
      groups.innerHTML = html;
    }
  }


  public async getLibraryItems(): Promise<IFileInfo[]> {
    //this.properties.libraryUrl = "https://maximusunitedkingdom.sharepoint.com/asm_dc/policies/";
    const list : Element | null = this.domElement.querySelector("#listContainer");
    let html : string = "";

    try {
      const items = await this.sp.web.lists.getByTitle("Documents").items();
      items.forEach((item:any)=>{
        html+=`<li>Item Title: ${item.Title}</li>`;
      })
      if(list) {
        list.innerHTML = html;
      }  
      return items;
    } catch (error) {
      console.error("Error fetching library items:", error);
      throw error;
    }
  }


  public render(): void {

    this.properties.userEmail = this.context.pageContext.user.email;
    //const emailDomain = this.properties.userEmail.split('@')[1];
    const graphContainer : Element | null = this.domElement.querySelector("#graphContainer");
    const spContainer : Element | null = this.domElement.querySelector("#spContainer");

    this.domElement.innerHTML = `
    <section class="${styles.testUserAccess} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
      </div>
      <div class="row">
        <h4>Email : ${this.properties.userEmail}</h4>
        <div class="col">
          <h4><u>User Groups - using Graph API</u></h4>
          <ul id="graphContainer"></ul>
        </div>
        <div class="col">
          <h4><u>User Groups - using SharePoint API</u></h4>
          <ul id="spContainer"></ul>
      </div>
      <div class="row" id="listContainer"/>
    </section>`;

    this.getGroupsGraph() //this.context.pageContext.user.loginName)
      .then(groupNames => this.renderGraphGroups(groupNames))
      .catch(error => {
        console.error("Error rendering user groups from GRAPH API:", error);
        if(graphContainer){
          graphContainer.innerHTML = `Error rendering user groups from GRAPH API: ${error}`;      
        }
      });
    
    this.getGroupsSP()
      .then(groups => {
        const groupsList = groups.map(group => `<li>Group Title: ${escape(group.Title)}</li>`).join('');
        if(spContainer){
          spContainer.innerHTML = groupsList || '<li> No Groups Found </li>';
        }        
      })
    
    //this.getLibraryItems()
  }

  public async onInit(): Promise<void> {
    await super.onInit();
    this.sp = spfi().using(spSPFx(this.context)).using(PnPLogging(LogLevel.Warning));  
    this.graph = graphfi().using(graphSPFx(this.context)).using(PnPLogging(LogLevel.Warning));  

    SPComponentLoader.loadCss("https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/css/bootstrap.min.css");
    SPComponentLoader.loadCss("https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.3/font/bootstrap-icons.css");

//    return super.onInit().then(_ => {

//    })

    return this._getEnvironmentMessage().then(message => {
      //this._environmentMessage = message;
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

    this._isDarkTheme = !!currentTheme.isInverted;
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
