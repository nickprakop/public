import { ServiceKey, ServiceScope } from '@microsoft/sp-core-library';
import { BaseService } from './BaseService';
import { IList } from '@pnp/sp/lists';
import { IFolder, IFolderInfo } from '@pnp/sp/folders';
import { IResourcePath } from '@pnp/sp';

export interface IFolderService {
    getFolder: (title: String) => Promise<IFolderInfo | null>;
    createFolder: (title: String) => Promise<ICreatingFolderResult>;
    buildFolderUrl: (folderNama: string) => string;
    initialize: (libraryTitle: string) => Promise<void>;
}

export const Folder_ServiceKey = 'FolderService';

export interface ICreatingFolderResult {
    isNewCreated: boolean;
    folder: IFolderInfo | null;
}

export class FolderService extends BaseService implements IFolderService {
    library: IList;
    rootFolder: IFolderInfo;

    public static readonly ServiceInstanceKey: ServiceKey<IFolderService> = ServiceKey.create(Folder_ServiceKey, FolderService);

    constructor(serviceScope: ServiceScope) {
        super(serviceScope);
    }

    async initialize(libraryTitle: string) {
        this.library = this.sp.web.lists.getByTitle(libraryTitle);
        this.rootFolder = (await this.library.rootFolder());
    }

    buildFolderUrl(folderName: string) {
        return `${this.rootFolder.ServerRelativeUrl}/${folderName}/`;
    }


    async getFolder(title: String): Promise<IFolderInfo | null> {
        const folders: IFolderInfo[] = await this.library.rootFolder.folders.filter(`Name eq '${title}'`)();
        if (folders.length > 0) {
            return folders[0];
        }
        return null;
    }

    async createFolder(title: string): Promise<ICreatingFolderResult> {
        let result: ICreatingFolderResult = { isNewCreated: false, folder: null };
        try {
            void await this.library.rootFolder.addSubFolderUsingPath(title);
            result.isNewCreated = true;
        } catch (e) {
            void this.logError(e);
        }
        result.folder = await this.getFolder(title);
        return result;
    }
}