export interface ITaskItem {
    listName: string,
    title: string,
    description?: string,
    important: boolean,
    // inferred from due date
    today: boolean,
    id: string,
    // useful for viva connections (works like an index)
    order?: number,
    dueDate?: string,
    overDueDays?: string,
    createdDate: Date,
    fromMail: boolean,
    platform: TaskPlatform
    logoUrl: PlatformLogo
}

export enum TaskPlatform {
    ToDo = "todo",
    AzureDevOps = "ado"
}

export enum PlatformLogo {
    ToDo = "https://reckittstorage.blob.core.windows.net/viva-connections-icons/mstodo.svg",
    AzureDevOps = "ado"
}