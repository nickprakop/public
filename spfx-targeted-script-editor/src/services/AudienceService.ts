export default class AudienceService {

  webAbsoluteUrl: string;
  
  constructor(webAbsoluteUrl: string) {
    this.webAbsoluteUrl = webAbsoluteUrl;
  }

  public async CheckAudiences(targetedGroups): Promise<boolean> {
    let checkResult = false;
    if (targetedGroups?.length === 0) {
      checkResult = true;
    } else {

      const userGroups = await this.getCurrentUserGroups()
      const targetedGroupNames = targetedGroups.map(gr => gr.login);
      const userInTargetGroups = targetedGroupNames.filter(gr => userGroups.indexOf(gr) > 0);
      if (userInTargetGroups.length > 0) {
        checkResult = true;
      } else {
        let proms: any[] = [];
        targetedGroups.map((item) => {
          proms.push(this.tryGetGroupMembers(item.fullName));
        });
        const results = await Promise.all(proms);
        checkResult = results.some(result => result === true);
      }

    }
    return checkResult;
  }
  /*check if user is a member of the group, using SP rest
  */
  public async isMember(groupName: string, userId: string): Promise<boolean> {
    const reqUrl = this.webAbsoluteUrl + "/_api/web/sitegroups/getByName('" + groupName + "')/Users?$filter=Id eq " + userId;
    try {
      const response = await fetch(`${reqUrl}`, {
        method: "GET",
        headers: {
          "Accept": "application/json;odata=verbose",
          "odata-version": ""
        }
      });

      if (response.ok) {
        return true;
      }

    } catch (error) {
      console.log(error);
    }
    return false;
  }

  public async tryGetGroupMembers(groupName: string): Promise<boolean> {
    const reqUrl = this.webAbsoluteUrl + "/_api/web/sitegroups/getByName('" + groupName + "')/Users?$top=1";
    try {
      const response = await fetch(reqUrl, {
        method: "GET",
        headers: {
          "Accept": "application/json;odata=verbose",
          "odata-version": ""
        }
      });

      if (response.status === 403 || !response.ok) {
        return false;
      }

      return true;
    } catch (error) {
      console.log(error);
    }
    return false;
  }

  public async getCurrentUserGroups(): Promise<string[]> {
    let groups = [];

    const reqUrl = this.webAbsoluteUrl + "/_api/web/currentuser/groups";
    try {
      const response = await fetch(`${reqUrl}`, {
        method: "GET",
        headers: {
          "Accept": "application/json;odata=verbose",
          "odata-version": ""
        }
      });

      if (response.ok) {
        const data = await response.json();
        groups = data.d.results.map(gr => gr.LoginName);
      }
    } catch (error) {
      console.log(error);
    }

    return groups;
  }
}