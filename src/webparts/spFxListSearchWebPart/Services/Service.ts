import { getSP } from "../pnpjsConfig";
import { SPFI } from "@pnp/sp";

import "@pnp/sp/webs";
import "@pnp/sp/lists/web";
import "@pnp/sp/fields";
import "@pnp/common";

export class SPService {
    private _sp : SPFI;

    constructor() {
        this._sp = getSP();
    }

    //----------Get fields/columns of selected List--------------------
    public async getFields(selectedList : string) : Promise<any> {
        try {
            const spList : any = this._sp.web.lists.getById(selectedList);
            const allFields : any[] = await spList.fields();

            return allFields;
        } catch (error) {
            Promise.reject(error)
        }
    }
}