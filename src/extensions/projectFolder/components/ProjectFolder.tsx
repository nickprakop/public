import { Link } from '@fluentui/react';
import { FieldCustomizerContext, ListItemAccessor } from '@microsoft/sp-listview-extensibility';
import { Project } from '@src/app/Project';
import { ProjectStatus } from '@src/app/ProjectStatus';
import { FolderService, ICreatingFolderResult, IFolderService } from '@src/services/FolderService';
import * as React from 'react';
import { FC, useMemo } from 'react';


export interface IProjectFolderProps {
  itemAccessor: ListItemAccessor;
  context: FieldCustomizerContext;
  libraryTitle: string;
}

export const ProjectFolder: FC<IProjectFolderProps> = (props) => {
  const { itemAccessor, context, libraryTitle } = props;
  const [folderService, setFolderService] = React.useState<IFolderService | null>(null);
  const [buttonCreateFolderText, setButtonCreateFolderText] = React.useState("Create Space");
  const [buttonDisabled, setButtonDisabled] = React.useState(true);
  const [project, setProject] = React.useState<Project>(new Project());

  React.useEffect(() => {
    setProject(project => ({
      ...project,
      id: itemAccessor.getValueByName("ID"),
      title: itemAccessor.getValueByName("Title"),
      status: itemAccessor.getValueByName("ProjectStatus"),
      folderName: project.folderName
    }));
  }, [itemAccessor.getValueByName("ProjectStatus")]);

  React.useEffect(() => {    
    context.serviceScope.whenFinished(async () => {
      let service = context.serviceScope.consume<IFolderService>(FolderService.ServiceInstanceKey);      
      await service.initialize(libraryTitle);
      setFolderService(service);
      setButtonDisabled(false);
    });
  }, [context]);

  const createFolder = async (folderName: string) => {
    setButtonCreateFolderText("Working on it...");
    setButtonDisabled(true);
    const creatingFolderResult: ICreatingFolderResult | undefined = await folderService?.createFolder(folderName);
    let folderUrl = creatingFolderResult?.folder?.ServerRelativeUrl;
    if (!creatingFolderResult?.isNewCreated) {
      alert("Folder already exists");
    }
    setButtonDisabled(false);
    window.open(folderUrl, '_blank');
  }

  const element: JSX.Element = useMemo(() => {
    switch (project.status) {
      case ProjectStatus.Planned:
        return <Link disabled={buttonDisabled} type="Link"
          onClick={(e) => {
            e.stopPropagation(); void createFolder(project.folderName());
          }}
        >{buttonCreateFolderText}</Link>
      case ProjectStatus.InProgress:
      case ProjectStatus.Published:
        return <Link disabled={buttonDisabled} onClick={() => { window.open(folderService?.buildFolderUrl(project.folderName()), '_blank'); }} target="_blank">Open Space</Link>
      case ProjectStatus.Rejected:
        return <></>

      default:
        return <>{project.status}</>
    }
  }, [project.status, buttonDisabled, buttonCreateFolderText]);

  return element;
}


