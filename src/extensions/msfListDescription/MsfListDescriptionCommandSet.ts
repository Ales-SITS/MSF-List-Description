import {
  BaseListViewCommandSet,
  ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/items";


export interface IMsfListDescriptionCommandSetProperties {

  sampleTextOne: string;
  sampleTextTwo: string;
}


export default class MsfListDescriptionCommandSet extends BaseListViewCommandSet<IMsfListDescriptionCommandSetProperties> {
  public onInit(): Promise<void> {
    console.log("Init run")

    const libhead = document.querySelector(".od-ItemsScopeItemContent-header")
    const sp = spfi().using(SPFx(this.context));
    const listTitle: string =  `${this.context.pageContext.list.title}`
  
    sp.web.lists.getByTitle(listTitle).select("Description")().then((result) => {
      console.log("rendering...")
      const old = document.getElementById("spfx_des")
      old === null? "" : old.remove()
      const description:string = result.Description;
      description.includes('<script') ? "" :
      description.startsWith('description=true:') ? 
      libhead.insertAdjacentHTML('afterend',`
      <div id="spfx_des" style="
      min-height: 45px;
      width: 100%;
      margin: 0 10px;
      padding: 0 20px;
      overflow: hidden;
      position: sticky;
      z-index: 1;
      top:5px
      ">${description.replace('description=true:','')}
      </div> 
      `) 
      : libhead.insertAdjacentHTML('beforeend',``);
    
      console.log("rendering finished...");
    }).catch((error) => {
      console.log(error);
    });

    this.context.listView.listViewStateChangedEvent.add(this, this.onListViewUpdatedv2);
    //this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

    return Promise.resolve();
  }

  public onListViewUpdatedv2(args: ListViewStateChangedEventArgs): void {
    console.log("Update run")
    document.getElementById("spfx_des").remove()
    const libhead = document.querySelector(".od-ItemsScopeItemContent-header")
    const sp = spfi().using(SPFx(this.context));
    const listTitle: string =  `${this.context.pageContext.list.title}`
  
    sp.web.lists.getByTitle(listTitle).select("Description")().then((result) => {
      console.log("rendering updated...")
      const old = document.getElementById("spfx_des")
      old === null? "" : old.remove()
      const description:string = result.Description;
      description.includes('<script') ? "" :
      description.startsWith('description=true:') ? 
      libhead.insertAdjacentHTML('afterend',`
      <div id="spfx_des" style="
      min-height: 45px;
      width: 100%;
      margin: 0 10px;
      padding: 0 20px;
      overflow: hidden;
      position: sticky;
      z-index: 1;
      top:5px
      ">${description.replace('description=true:','')}
      </div> 
      `) 
      : libhead.insertAdjacentHTML('beforeend',``);
    
      //console.log(description);
    }).catch((error) => {
      console.log(error);
    });
  }



}
