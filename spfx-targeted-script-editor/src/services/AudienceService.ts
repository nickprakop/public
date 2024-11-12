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

      const errors: string[] = [];

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
    const response = await fetch(`${reqUrl}`, {
      method: "GET",
      headers: {
        "Accept": "application/json;odata=verbose",
        "odata-version": ""
      }
    });

    if (!response.ok) {
      return false;
    }

    return true;
  }

  public async tryGetGroupMembers(groupName: string): Promise<boolean> {
    const reqUrl = this.webAbsoluteUrl + "/_api/web/sitegroups/getByName('" + groupName + "')/Users?$top=1"
    const response = await fetch(`${reqUrl}`, {
      method: "GET",
      headers: {
        "Accept": "application/json;odata=verbose",
        "odata-version": ""
      }
    });
    if (!response.ok) {
      return false;
    }
    const data = await response.json();
    if (data.d.results.length > 0) {
      return true;
    }
    return false;
  }

  public async getCurrentUserGroups(): Promise<string[]> {
    const reqUrl = this.webAbsoluteUrl + "/_api/web/currentuser/groups";
    const response = await fetch(`${reqUrl}`, {
      method: "GET",
      headers: {
        "Accept": "application/json;odata=verbose",
        "odata-version": ""
      }
    });
    if (!response.ok) {
      return [];
    }
    const data = await response.json();
    let groups = data.d.results.map(gr => gr.LoginName);
    return groups;
  }
}