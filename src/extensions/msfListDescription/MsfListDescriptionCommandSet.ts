import {
  BaseListViewCommandSet
} from '@microsoft/sp-listview-extensibility';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/items";


export interface IMsfListDescriptionCommandSetProperties {
}

export default class MsfListDescriptionCommandSet extends BaseListViewCommandSet<IMsfListDescriptionCommandSetProperties> {
  public onInit(): Promise<void> {
    
    let pageUrl = this.context.pageContext.list.serverRelativeUrl
    const libhead = document.querySelector(".od-ItemsScopeItemContent-header")
    const sp = spfi().using(SPFx(this.context));
    const listTitle: string =  `${this.context.pageContext.list.title}`
  
    sp.web.lists.getByTitle(listTitle).select("Description")().then((result) => {
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
      top:5px;
      left:5px;
      z-index: 1;
      ">${description.replace('description=true:','')}
      </div> 
      `) 
      : libhead.insertAdjacentHTML('beforeend',``);
    }).catch((error) => {
      console.log(error);
    });

    setInterval(()=>{
       if(pageUrl!==this.context.pageContext.list.serverRelativeUrl) {  


    const libhead = document.querySelector(".od-ItemsScopeItemContent-header")
    const sp = spfi().using(SPFx(this.context));
    const listTitle: string =  `${this.context.pageContext.list.title}`
  
    sp.web.lists.getByTitle(listTitle).select("Description")().then((result) => {
     
      const old = document.getElementById("spfx_des")
      old === null? "" : old.remove()
      const description:string = result.Description;
      description.includes('<script') ? "" :
      description.startsWith('description=true:') ? 
      libhead.insertAdjacentHTML('afterend',`
      <div id="spfx_des" class="od-ItemsScopeList-content-sticky" style="
      min-height: 45px;
      width: 100%;
      margin: 0 10px;
      padding: 0 20px;
      overflow: hidden;
      position: sticky;
      top:5px;
      left:5px;
      z-index: 1;
      ">${description.replace('description=true:','')}
      </div> 
      `) 
      : libhead.insertAdjacentHTML('beforeend',``);
    }).catch((error) => {
      console.log(error);
    });
    
      }
  },4000)

    return Promise.resolve();
  }

}
