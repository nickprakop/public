import * as React from 'react';
import { IPropertyFieldGroupOrPerson } from '@pnp/spfx-property-controls';
import { BaseComponentContext } from '@microsoft/sp-component-base';
import spservices from '../../../../services/spservices';
import { PageContext } from '@microsoft/sp-page-context';
import { SPComponentLoader } from '@microsoft/sp-loader';

export interface IScriptExecutorProps {
  scriptBody: string;
  targetedGroups: IPropertyFieldGroupOrPerson[];
  pageContext: PageContext,
  removePadding: boolean;
  unqiueId: string;
  spPageContextInfo: boolean;
  teamsContext: boolean;
}

export interface IScriptExecutorState {
  canView: boolean;
}

export default class ScriptExecutor extends React.Component<IScriptExecutorProps, any> {

  element: HTMLElement;

  constructor(props: IScriptExecutorProps) {
    super(props);

    this.state = {
      canView: false
    } as IScriptExecutorState;

  }
  public componentDidMount(): void {
    this.element = document.createElement('div');
    this.element.innerHTML = this.props.scriptBody;
    //setting the state whether user has permission to view webpart
    if (this.props.targetedGroups?.length > 0) {
      this.checkUserCanViewWebpart();
    } else {
      this.setState({ canView: true });
    }
  }

  public checkUserCanViewWebpart(): void {
    const self = this;
    let proms: any[] = [];
    const errors: string[] = [];
    const _sv = new spservices();
    self.props.targetedGroups.map((item) => {
      proms.push(_sv.isMember(item.fullName, self.props.pageContext.legacyPageContext[`userId`], self.props.pageContext.site.absoluteUrl));
    });
    Promise.race(
      proms.map(p => {
        return p.catch(err => {
          errors.push(err);
          if (errors.length >= proms.length) throw errors;
          return new Promise(() => { });
        });
      })).then(val => {
        this.setState({ canView: true }); //atleast one promise resolved
      });
  }

  public render(): JSX.Element {
    return (
      <div>
        {
          this.state.canView ?
            this.executeScript(this.element)
            : ``
        }
      </div>
    );
  }

  private evalScript(elem) {
    const data = (elem.text || elem.textContent || elem.innerHTML || "");
    const headTag = document.getElementsByTagName("head")[0] || document.documentElement;
    const scriptTag: HTMLScriptElement = document.createElement("script");

    for (let i = 0; i < elem.attributes.length; i++) {
      const attr = elem.attributes[i];
      // Copies all attributes in case of loaded script relies on the tag attributes
      if (attr.name.toLowerCase() === "onload") continue; // onload handled after loading with SPComponentLoader
      scriptTag.setAttribute(attr.name, attr.value);
    }

    // set a bogus type to avoid browser loading the script, as it's loaded with SPComponentLoader
    scriptTag.type = scriptTag.src?.length > 0 ? "pnp" : "text/javascript";
    // Ensure proper setting and adding id used in cleanup on reload
    scriptTag.setAttribute("pnpname", this.props.unqiueId);

    try {
      // doesn't work on ie...
      scriptTag.appendChild(document.createTextNode(data));
    } catch (e) {
      // IE has funky script nodes
      scriptTag.text = data;
    }

    headTag.insertBefore(scriptTag, headTag.firstChild);
  }

  private async executeScript(element: HTMLElement) {
    // clean up added script tags in case of smart re-load        
    const headTag = document.getElementsByTagName("head")[0] || document.documentElement;
    let scriptTags = headTag.getElementsByTagName("script");
    for (let i = 0; i < scriptTags.length; i++) {
      const scriptTag = scriptTags[i];
      if (scriptTag.hasAttribute("pnpname") && scriptTag.attributes["pnpname"].value == this.props.unqiueId) {
        headTag.removeChild(scriptTag);
      }
    }

    if (this.props.spPageContextInfo && !window["_spPageContextInfo"]) {
      window["_spPageContextInfo"] = this.context.pageContext.legacyPageContext;
    }

    if (this.props.teamsContext && !window["_teamsContexInfo"]) {
      window["_teamsContexInfo"] = this.context.sdks.microsoftTeams?.context;
    }

    // Define global name to tack scripts on in case script to be loaded is not AMD/UMD
    window['ScriptGlobal'] = {};

    // main section of function
    const scripts: any[] = [];
    const children_nodes = element.getElementsByTagName("script");

    for (let i = 0; children_nodes[i]; i++) {
      const child: any = children_nodes[i];
      if (!child.type || child.type.toLowerCase() === "text/javascript") {
        scripts.push(child);
      }
    }

    const urls: string[] = [];
    const onLoads: any[] = [];
    for (let i = 0; scripts[i]; i++) {
      const scriptTag: HTMLScriptElement = scripts[i];
      if (scriptTag.src && scriptTag.src.length > 0) {
        urls.push(scriptTag.src);
      }
      if (scriptTag.onload && scriptTag.onload.length > 0) {
        onLoads.push(scriptTag.onload);
      }
    }

    let oldamd = null;
    if (window["define"] && window["define"].amd) {
      oldamd = window["define"].amd;
      window["define"].amd = null;
    }

    for (let i = 0; i < urls.length; i++) {
      try {
        let scriptUrl = urls[i];
        // Add unique param to force load on each run to overcome smart navigation in the browser as needed
        const prefix = scriptUrl.indexOf('?') === -1 ? '?' : '&';
        scriptUrl += prefix + 'pnp=' + new Date().getTime();
        await SPComponentLoader.loadScript(scriptUrl, { globalExportsName: "ScriptGlobal" });
      } catch (error) {
        if (console.error) {
          console.error(error);
        }
      }
    }
    if (oldamd) {
      window["define"].amd = oldamd;
    }

    for (let i = 0; scripts[i]; i++) {
      const scriptTag = scripts[i];
      if (scriptTag.parentNode) { scriptTag.parentNode.removeChild(scriptTag); }
      this.evalScript(scripts[i]);
    }
    // execute any onload people have added
    for (let i = 0; onLoads[i]; i++) {
      onLoads[i]();
    }
  }
  // public render(): React.ReactElement<ScriptExecutorProps> {
  //     if (this.properties.removePadding) {
  //         let element: HTMLElement | null = this.domElement.parentElement;
  //         // check up to 5 levels up for padding and exit once found
  //         for (let i = 0; i < 5 && element; i++) {
  //           const style = window.getComputedStyle(element);
  //           const hasPadding = style.paddingTop !== "0px";
  //           if (hasPadding) {
  //             element.style.paddingTop = "0px";
  //             element.style.paddingBottom = "0px";
  //             element.style.marginTop = "0px";
  //             element.style.marginBottom = "0px";
  //           }
  //           element = element.parentElement;
  //         }
  //       }

  //       ReactDom.unmountComponentAtNode(this.domElement);
  //       this.domElement.innerHTML = this.properties.scriptBody;
  //       this.executeScript(this.domElement);
  // }
}
