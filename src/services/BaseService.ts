import { ServiceScope } from "@microsoft/sp-core-library";
import { ISPFXContext, SPFI, SPFx, spfi } from '@pnp/sp/presets/all';
import { PageContext } from "@microsoft/sp-page-context";
import { HttpRequestError } from "@pnp/queryable";

export abstract class BaseService {
    sp: SPFI;

    constructor(serviceScope: ServiceScope) {
        serviceScope.whenFinished(() => {
            const pageContext = serviceScope.consume(PageContext.serviceKey)
            this.sp = spfi().using(SPFx({ pageContext }));
        })
    }

    async logError(e: HttpRequestError): Promise<void> {
        if (e?.isHttpRequestError) {

            // we can read the json from the response
            const json = await (<HttpRequestError>e).response.json();

            // if we have a value property we can show it
            console.log(typeof json["odata.error"] === "object" ? json["odata.error"].message.value : e.message);

            // add of course you have access to the other properties and can make choices on how to act
            if ((<HttpRequestError>e).status === 404) {
                console.error((<HttpRequestError>e).statusText);
                // maybe create the resource, or redirect, or fallback to a secondary data source
                // just ideas, handle any of the status codes uniquely as needed
            }

        } else {
            // not an HttpRequestError so we just log message
            console.log(e.message);
        }
    }
}