export interface ITaskItem {
    listName?: string,
    title: string,
    description?: string,
    important?: boolean,
    // inferred from due date
    today: boolean,
    id: string,
    // useful for viva connections (works like an index)
    order?: number,
    dueDate?: string,
    overDueDays?: string,
    createdDate: Date,
    createdDateString: string,
    fromMail?: boolean,
    platform: TaskPlatform
    logoUrl: PlatformLogo
    source?: string
    deepLinkUrl?: string
    percentComplete?: number
    listId?: string
}

export enum TaskPlatform {
    ToDo = "todo",
    AzureDevOps = "ado",
    Planner = "planner",
}

export enum PlatformLogo {
    ToDo = "https://reckittstorage.blob.core.windows.net/viva-connections-icons/mstodo.svg",
    AzureDevOps = "https://reckittstorage.blob.core.windows.net/viva-connections-icons/azuredevops.svg",
    Planner = "https://reckittstorage.blob.core.windows.net/viva-connections-icons/planner.png",

}