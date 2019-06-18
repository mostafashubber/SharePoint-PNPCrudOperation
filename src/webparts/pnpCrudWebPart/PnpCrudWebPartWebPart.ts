import * as pnp from 'sp-pnp-js'; 
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './PnpCrudWebPartWebPart.module.scss';
import * as strings from 'PnpCrudWebPartWebPartStrings';


export interface ISPList {
  ID: string;
  Title: string;
  Name: string;
  Country: string;
  City: string;
} 
export interface IPnpCrudWebPartWebPartProps {
  description: string;
}

export default class PnpCrudWebPartWebPart extends BaseClientSideWebPart<IPnpCrudWebPartWebPartProps> {



  private AddEventListeners() : void{
    document.getElementById('AddItem').addEventListener('click',()=>this.AddItem());
    document.getElementById('UpdateItem').addEventListener('click',()=>this.UpdateItem());
    document.getElementById('DeleteItem').addEventListener('click',()=>this.DeleteItem());
   }
   
    private _getListData(): Promise<ISPList[]> {
    return pnp.sp.web.lists.getByTitle("EmployeeList").items.get().then((response) => {
      
       return response;
     });
        
   }
   
    private getListData(): void {
      
       this._getListData()
         .then((response) => {
           this._renderList(response);
         });
   }

   private _renderList(items: ISPList[]): void {
    let html: string = '<table>';
    html += `<th>ID</th><th>Title</th><th>Name</th><th>Country</th><th>City</th>`;
    items.forEach((item: ISPList) => {
      html += `
           <tr>
          <td>${item.ID}</td>
          <td>${item.Title}</td>
          <td>${item.Name}</td>
          <td>${item.Country}</td>
          <td>${item.City}</td>
          </tr>
          `; 
    });
    html += `</table>`;
    const listContainer: Element = this.domElement.querySelector('#spGetListItems');
    listContainer.innerHTML = html;
  }


  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.pnpCrudWebPart }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
          <h1>Employee List</h1>
          <form >
          <br>
          <div data-role="header">
             <h2>Add SharePoint List Items</h2>
          </div>
          <br>
          <div data-role="main" class="ui-content">
             <div >
                <input id="Title" type ="text" placeholder="Title" />
                <input id="Name" type ="text" placeholder="Name"  />
                <input id="Country" type ="text" placeholder="Country" />
                <input id="City" type ="text" placeholder="City" />
             </div>
             <div></br></div>
             <div >
                <button id="AddItem"  type="submit" >Add</button>
             </div>
             
          </div>
          <br>
          <div data-role="header">
             <h2>Update/Delete SharePoint List Items</h2>
          </div>
          <br>
          <div data-role="main" class="ui-content">
             <div >
                <input id="ID"  type="text" placeholder="ID"  />
             </div>
             <div></br></div>
             <div >
                <button id="UpdateItem" type="submit" >Update</button>
                &nbsp
                <button id="DeleteItem"  type="submit" >Delete</button>
             </div>
          </div>
       </form>
       <br>
       <br>
      <div id="spGetListItems"/>
      </div>
      </div>
      </div>`;

      this.getListData();
      this.AddEventListeners();
  }

  protected AddItem()
{  
   
     pnp.sp.web.lists.getByTitle('EmployeeList').items.add({    
     Title : document.getElementById('Title')["value"],
     Name : document.getElementById('Name')["value"],
     Country:document.getElementById('Country')["value"],
     City:document.getElementById('City')["value"]
    });
   
   
}

protected UpdateItem()
{  
    
    var id = document.getElementById('ID')["value"];
    pnp.sp.web.lists.getByTitle("EmployeeList").items.getById(id).update({
     Title : document.getElementById('Title')["value"],
     Name : document.getElementById('Name')["value"],
     Country:document.getElementById('Country')["value"],
     City:document.getElementById('City')["value"]
  });
 
}

 protected DeleteItem()
{  
     pnp.sp.web.lists.getByTitle("EmployeeList").items.getById(document.getElementById('ID')["value"]).delete();
     
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
