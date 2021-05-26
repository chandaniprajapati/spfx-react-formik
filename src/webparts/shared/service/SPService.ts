import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import { Web } from "@pnp/sp/webs";

export class SPService {

    private web;

    constructor(url: string) {
        this.web = Web(url);
    }

    public async createTask(listName: string, body: any) {
        try {
            let createdItem = await this.web.lists
                .getByTitle(listName)
                .items
                .add(body);
            return createdItem;
        }
        catch (err) {
            Promise.reject(err);
        }
    }
}