import * as $ from 'jquery';

export default class spservices {

    /*check if user is a member of the group, using SP rest
    */
    public async isMember(groupName: string, userId: string, webAbsoluteUrl): Promise<any> {
        let p = new Promise<any>((resolve, reject) => {
            $.ajax({
                url: webAbsoluteUrl + "/_api/web/sitegroups/getByName('" + groupName + "')/Users?$filter=Id eq " + userId,
                method: "GET",
                headers: { "Accept": "application/json; odata=verbose" },
                success: (data) => {
                    if (data.d.results[0] != undefined) {
                        resolve(true);
                    }
                    else {
                        reject(false);
                    }
                },
                error: (error) => {
                    reject(false);
                },
            });
        });
        return p;
    }

    public async tryGetGroupMembers(groupName: string, webAbsoluteUrl): Promise<any> {
        let p = new Promise<any>((resolve, reject) => {
            $.ajax({
                url: webAbsoluteUrl + "/_api/web/sitegroups/getByName('" + groupName + "')/Users",
                method: "GET",
                headers: { "Accept": "application/json; odata=verbose" },
                success: (data) => {
                    if (data.d.results.length > 0) {
                        resolve(true);
                    }
                    else {
                        reject(false);
                    }
                },
                error: (error) => {
                    reject(false);
                },
            });
        });
        return p;
    }

    public async getCurrentUserGroups(webAbsoluteUrl): Promise<any> {
        let p = new Promise<any>((resolve, reject) => {
            $.ajax({
                url: webAbsoluteUrl + "/_api/web/currentuser/groups",
                method: "GET",
                headers: { "Accept": "application/json; odata=verbose" },
                success: (data) => {
                    let groups = data.d.results.map(gr => gr.LoginName);
                    resolve(groups);
                },
                error: (error) => {
                    reject(error);
                },
            });
        });
        return p;
    }
}