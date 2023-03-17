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
    createdDate: Date,
    fromMail: boolean,
    platform: TaskPlatform
}

export enum TaskPlatform {
    ToDo = "todo",
    AzureDevOps = "ado"
}