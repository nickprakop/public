import * as React from 'react';
import * as ReactDOM from 'react-dom';

import {
  BaseFieldCustomizer,
  type IFieldCustomizerCellEventParameters
} from '@microsoft/sp-listview-extensibility';
import { IProjectFolderProps, ProjectFolder } from './components/ProjectFolder';
import { Folder_ServiceKey } from '@src/services/FolderService';
/**
 * If your field customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IProjectFolderFieldCustomizerProperties {
  // This is an example; replace with your own property
  listNameWithFolders: string;
}

const LOG_SOURCE: string = 'ProjectFolderFieldCustomizer';

export default class ProjectFolderFieldCustomizer extends BaseFieldCustomizer<IProjectFolderFieldCustomizerProperties> {
  public async onInit(): Promise<void> {    
    await super.onInit();
  }

  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
    // Use this method to perform your custom cell rendering.
    const projectFolder: React.ReactElement = React.createElement(ProjectFolder,
      { itemAccessor: event.listItem, context: this.context, libraryTitle: this.properties.listNameWithFolders } as IProjectFolderProps);
    ReactDOM.render(projectFolder, event.domElement);
  }

  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    // This method should be used to free any resources that were allocated during rendering.
    // For example, if your onRenderCell() called ReactDOM.render(), then you should
    // call ReactDOM.unmountComponentAtNode() here.
    ReactDOM.unmountComponentAtNode(event.domElement);
    super.onDisposeCell(event);
  }
}
