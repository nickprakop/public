
/* eslint-disable */
import { Version } from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { IPropertyPaneConfiguration, IPropertyPaneField, PropertyPaneButton, PropertyPaneButtonType, PropertyPaneToggle } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { Placeholder } from '@pnp/spfx-controls-react';
import { IPropertyFieldGroupOrPerson } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import * as strings from 'TargetedScriptEditorWebPartStrings';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import spservices from '../../services/spservices';

export interface ITargetedScriptEditorWebPartProps {
  description: string;
  scriptBody: string;
  removePadding: boolean;
  spPageContextInfo: boolean;
  teamsContext: boolean;
  targetedGroups: IPropertyFieldGroupOrPerson[];
}

export interface ITargetedScriptEditorWebPartState {
  exucuteScript: boolean;
}

export default class TargetedScriptEditorWebPart extends BaseClientSideWebPart<ITargetedScriptEditorWebPartProps> {
  public _scriptBodyEditorPanel;
  public _spPeoplePicker;

  public _unqiueId;


  constructor() {
    super();
    this.scriptUpdate = this.scriptUpdate.bind(this);
    this.onTargetedGroupsChanged = this.onTargetedGroupsChanged.bind(this);
  }

  public scriptUpdate(_property: string, _oldVal: string, newVal: string) {
    this.properties.scriptBody = newVal;
    this._scriptBodyEditorPanel.initialValue = newVal;

  }

  protected _onConfigure = () => {
    // Context of the web part
    this.context.propertyPane.open();
  }

  public render(): void {
    if (this.properties.scriptBody?.length > 0) {
      //if (this.displayMode === DisplayMode.Read) {
      // const element: React.ReactElement<ScriptExecutorProps> = React.createElement(
      //   ScriptExecutor,
      //   {
      //     context: this.context,
      //     targetedGroups: this.properties.targetedGroups,
      //     scriptBody: this.properties.scriptBody,
      //     removePadding: this.properties.removePadding
      //   }
      // );
      // ReactDom.render(element, this.domElement);

      if (this.properties.targetedGroups?.length > 0) {
        let proms: any[] = [];
        const errors: string[] = [];
        const _sv = new spservices();
        this.properties.targetedGroups.map((item) => {
          proms.push(_sv.isMember(item.fullName, this.context.pageContext.legacyPageContext[`userId`], this.context.pageContext.site.absoluteUrl));
        });
        Promise.race(
          proms.map(p => {
            return p.catch(err => {
              errors.push(err);
              if (errors.length >= proms.length) throw errors;
              return new Promise(() => { });
            });
          })).then(val => {
            ReactDom.unmountComponentAtNode(this.domElement);
            this.domElement.innerHTML = this.properties.scriptBody;
            this.executeScript(this.domElement);
          });
      } else {
        ReactDom.unmountComponentAtNode(this.domElement);
        this.domElement.innerHTML = this.properties.scriptBody;
        this.executeScript(this.domElement);
      }
      // } else {
      //   this.renderEditor();
      // }
    } else {
      const placeHolderElement = React.createElement(Placeholder, {
        iconName: "Edit",
        iconText: "Configure your web part",
        description: "Please configure the web part.",
        buttonLabel: "Configure",
        onConfigure: this._onConfigure,
      });
      ReactDom.render(placeHolderElement, this.domElement);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private saveClick() {
    this.context.propertyPane.refresh();
    this.onDispose();
  }

  private onTargetedGroupsChanged(propertyPath: string, oldValue: any, newValue: any) {
    this._spPeoplePicker.properties.initialData = newValue;
    this.properties.targetedGroups = newValue;
  }

  protected async loadPropertyPaneResources(): Promise<void> {
    //import { PropertyFieldCodeEditor, PropertyFieldCodeEditorLanguages } from '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor';
    const editorProp = await import(
      /* webpackChunkName: 'scripteditor' */
      '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor'
    );

    // import { PropertyFieldPeoplePicker, IPropertyFieldGroupOrPerson, PrincipalType } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
    const peoplePicker = await import(
      /* webpackChunkName: 'scripteditor' */
      '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker'
    );

    this._scriptBodyEditorPanel = editorProp.PropertyFieldCodeEditor('scriptBody', {
      label: 'Edit Code',
      panelTitle: 'Edit Code',
      initialValue: this.properties.scriptBody,
      onPropertyChange: this.scriptUpdate,
      properties: this.properties,
      disabled: false,
      key: 'codeEditorFieldId',
      language: editorProp.PropertyFieldCodeEditorLanguages.JavaScript
    });

    this._spPeoplePicker = peoplePicker.PropertyFieldPeoplePicker('targetedGroups', {
      label: 'Target Audience',
      initialData: this.properties.targetedGroups,
      allowDuplicate: false,
      principalType: [peoplePicker.PrincipalType.SharePoint],
      onPropertyChange: this.onTargetedGroupsChanged,
      context: this.context as any,
      properties: this.properties,
      onGetErrorMessage: undefined,
      deferredValidationTime: 0,
      key: 'groupsFieldId'
    })

  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

    let webPartOptions: IPropertyPaneField<any>[] = [
      PropertyPaneToggle("spPageContextInfo", {
        label: "Enable classic _spPageContextInfo",
        checked: this.properties.spPageContextInfo,
        onText: "Enabled",
        offText: "Disabled"
      }),
      PropertyPaneToggle("removePadding", {
        label: "Remove top/bottom padding of web part container",
        checked: this.properties.removePadding,
        onText: "Remove padding",
        offText: "Keep padding"
      }),
      this._scriptBodyEditorPanel,
      this._spPeoplePicker,
      PropertyPaneButton("save", {
        text: "Save",
        buttonType: PropertyPaneButtonType.Primary,
        onClick: this.saveClick.bind(this)
      })
    ];

    if (this.context.sdks.microsoftTeams) {
      let config = PropertyPaneToggle("teamsContext", {
        label: "Enable teams context as _teamsContexInfo",
        checked: this.properties.teamsContext,
        onText: "Enabled",
        offText: "Disabled"
      });
      webPartOptions.push(config);
    }

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupFields: webPartOptions
            }
          ]
        }
      ]
    };
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
    scriptTag.setAttribute("pnpname", this._unqiueId);

    try {
      // doesn't work on ie...
      scriptTag.appendChild(document.createTextNode(data));
    } catch (e) {
      // IE has funky script nodes
      scriptTag.text = data;
    }

    headTag.insertBefore(scriptTag, headTag.firstChild);
  }

  // Finds and executes scripts in a newly added element's body.
  // Needed since innerHTML does not run scripts.
  //
  // Argument element is an element in the dom.
  private async executeScript(element: HTMLElement) {
    // clean up added script tags in case of smart re-load        
    const headTag = document.getElementsByTagName("head")[0] || document.documentElement;
    let scriptTags = headTag.getElementsByTagName("script");
    for (let i = 0; i < scriptTags.length; i++) {
      const scriptTag = scriptTags[i];
      if (scriptTag.hasAttribute("pnpname") && scriptTag.attributes["pnpname"].value == this._unqiueId) {
        headTag.removeChild(scriptTag);
      }
    }

    if (this.properties.spPageContextInfo && !window["_spPageContextInfo"]) {
      window["_spPageContextInfo"] = this.context.pageContext.legacyPageContext;
    }

    if (this.properties.teamsContext && !window["_teamsContexInfo"]) {
      window["_teamsContexInfo"] = this.context.sdks.microsoftTeams?.context;
    }

    // Define global name to tack scripts on in case script to be loaded is not AMD/UMD
    (<any>window).ScriptGlobal = {};

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
}
