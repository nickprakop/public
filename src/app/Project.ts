import { ProjectStatus } from "./ProjectStatus";


export class Project {
    status: ProjectStatus
    title: string;
    id: number;
    folderName(): string {
        return `ProjectSpace_${this.id}`;
    }
}