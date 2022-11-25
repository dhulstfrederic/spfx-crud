import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';
import {SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions} from '@microsoft/sp-http';


import styles from './CruddemoWebPart.module.scss';
import * as strings from 'CruddemoWebPartStrings';
import { ISoftwareListItem } from './ISoftwareListItem';

export interface ICruddemoWebPartProps {
  description: string;
}

export default class CruddemoWebPart extends BaseClientSideWebPart<ICruddemoWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.cruddemo} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
      </div>
      <div>
       ID: <input type='text' id='txtId'/><br/><br/>
       <input type='button' id='btnRead' value='Read details'/><br/><br/>

       Title: <input type='text' id='txtSoftwareTitle'/><br/><br/>

      
      <input type='button' id='btnAddListItem' value='Add new item'/><br/><br/>
      <input type='button' id='btnDelete' value='Delete'/><br/><br/>
      <input type='button' id='btnUpdate' value='Update'/><br/><br/>
      <input type='button' id='btnShowAll' value='Show all'/><br/><br/>
  </div>
  <div id="divStatus">
  </div
    </section>`;
    this.bindElements();
  }
  private bindElements(): void {
    this.domElement.querySelector("#btnAddListItem").addEventListener('click', ()=>{this.AddListItem()});
    this.domElement.querySelector("#btnUpdate").addEventListener('click', ()=>{this.UpdateListItem()});
    this.domElement.querySelector("#btnRead").addEventListener('click', ()=>{this.ReadListItem()});
    this.domElement.querySelector("#btnShowAll").addEventListener('click', ()=>{this.ReadAllListItems()});
  }

   private ReadListItem(): void {
    let id = (document.getElementById("txtId") as HTMLInputElement).value;
    this._getListItemById(id).then(listItem=> {
      const input = (document.getElementById("txtSoftwareTitle")  as HTMLInputElement).value = listItem.Title;
    }).catch(error=> {
      let message : Element = this.domElement.querySelector("#divStatus");
      message.innerHTML = "Read could not fetch details "+error.message;
    })
  }

  private _getListItemById(id: string) : Promise<ISoftwareListItem> {
    const url : string  = this.context.pageContext.web.absoluteUrl+"/_api/web/lists/getbytitle('SoftwareCatalog')/items?$filter=ID eq "+id;

    return new Promise<ISoftwareListItem>((resolve, reject) => { 
        this.context.spHttpClient.get(url,SPHttpClient.configurations.v1).then((response:SPHttpClientResponse)=>{
          return response.json();
        }).then((listItems:any) =>{
          const untypedItem: any = listItems.value[0];
          let listItem = untypedItem as ISoftwareListItem;
          resolve(listItem);
        });
      }); 
  }

  private ReadAllListItems(): void {
    let html : string = "<table><th>ID</th><th>Title</th>";
    this._getListItems().then(listItems=> {
      listItems.forEach(listItem => {
        html+=`<tr><td>${listItem.ID}</td><td>${listItem.Title}</td></tr>`
        
      });
      html+=`</table>`;
      const listContainer : Element = this.domElement.querySelector("#divStatus")
      listContainer.innerHTML = html;
    }).catch(error=> {
      let message : Element = this.domElement.querySelector("#divStatus");
      message.innerHTML = "Read could not fetch details "+error.message;
    })
  }

  
  private _getListItems() : Promise<ISoftwareListItem[]> {
    const url : string  = this.context.pageContext.web.absoluteUrl+"/_api/web/lists/getbytitle('SoftwareCatalog')/items";

    return new Promise<ISoftwareListItem[]>((resolve, reject) => { 
        this.context.spHttpClient.get(url,SPHttpClient.configurations.v1).then((response:SPHttpClientResponse)=>{
          return response.json();
        }).then(json =>{
          const untypedItem: any = json.value;
          let result = untypedItem as ISoftwareListItem[];
          resolve(result);
        });
    });
  }

  private AddListItem(): void {
    let txtSoftwareTitle = (document.getElementById("txtSoftwareTitle") as HTMLInputElement).value;
    const url : string  = this.context.pageContext.web.absoluteUrl+"/_api/web/lists/getbytitle('SoftwareCatalog')/items";

    const itemBody : any = {
      "Title" : txtSoftwareTitle
    }

    const SPHttpClientOptions: ISPHttpClientOptions= {
      "body" : JSON.stringify(itemBody)
    }

    this.context.spHttpClient.post(url,SPHttpClient.configurations.v1,SPHttpClientOptions)
    .then((response:SPHttpClientResponse)=>{
      let statusMessage : Element = this.domElement.querySelector("#divStatus");

      if (response.status == 201) {
      //  alert("A new item has been added");
        statusMessage.innerHTML = "List item has been created successfully";
      } else {
        //alert("Error message: "+response.status+" - "+ response.bodyUsed + " - "+response.statusText);
        statusMessage.innerHTML = "Error message: "+response.status+" - "+ response.bodyUsed + " - "+response.statusText;
      }
      this.clear();
    });    
  }

  private UpdateListItem(): void {
    let txtSoftwareTitle = (document.getElementById("txtSoftwareTitle") as HTMLInputElement).value;
    let id = (document.getElementById("txtId") as HTMLInputElement).value;

    const url : string  = this.context.pageContext.web.absoluteUrl+"/_api/web/lists/getbytitle('SoftwareCatalog')/items("+id+")";

    const itemBody : any = {
      "Title" : txtSoftwareTitle
    }
    const headers: any = {
      "X-HTTP-MEthod": "MERGE",
      "IF-MATCH": "*"
    }

    const SPHttpClientOptions: ISPHttpClientOptions= {
      "headers" : headers,
      "body" : JSON.stringify(itemBody)
    }

    this.context.spHttpClient.post(url,SPHttpClient.configurations.v1,SPHttpClientOptions)
    .then((response:SPHttpClientResponse)=>{
      let statusMessage : Element = this.domElement.querySelector("#divStatus");

      if (response.status == 204) {
      //  alert("A new item has been added");
        statusMessage.innerHTML = "List item has been updated successfully";
      } else {
        //alert("Error message: "+response.status+" - "+ response.bodyUsed + " - "+response.statusText);
        statusMessage.innerHTML = "Error message: "+response.status+" - "+ response.bodyUsed + " - "+response.statusText;
      }
      this.clear();
    });    

  }

  private clear(): void {
    (document.getElementById("txtSoftwareTitle") as HTMLInputElement).value = "";
  }

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }



  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
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
