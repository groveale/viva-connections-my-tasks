import { MSGraphClientV3 } from '@microsoft/sp-http'
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { ITaskItem, TaskPlatform } from '../models/ITask';

export enum WellKnownNames {
    FlaggedEmails = "flaggedEmails",
    Tasks = "defaultList"
}

export enum Importance {
    Normal = "normal",
    Important = "high"
}

export interface IGraphService {
    Init: (graphClient: MSGraphClientV3, loggedinUsersUPN: string) => void;    
    GetOutstandingTodoItemsInList: (list: MicrosoftGraph.TodoTaskList) => Promise<MicrosoftGraph.TodoTask[]>;
    GetUsersTaskLists: () => Promise<MicrosoftGraph.TodoTaskList[]>;
    GetOutStandingTaskFromToDo: () => Promise<ITaskItem[]>;
    GetITaskItemFromToDoItem: (list: MicrosoftGraph.TodoTaskList, todoItem: MicrosoftGraph.TodoTask) => ITaskItem

}

export class GraphService implements IGraphService {
    
    
    private graphClient: MSGraphClientV3;
    private upn: string
    private flaggedEmailListId: string
    
    public Init(graphClient: MSGraphClientV3, loggedinUsersUPN: string): void {
        this.graphClient = graphClient
        this.upn = loggedinUsersUPN
    }

    public async GetOutstandingTodoItemsInList(list: MicrosoftGraph.TodoTaskList): Promise<MicrosoftGraph.TodoTask[]> {
        if (this.graphClient === undefined){
            throw new Error('GraphService not initialized!')
        }

        // return tasks that are not completed in the list
        return await this.graphClient.api(`/me/todo/lists/${list.id}/tasks`)
                .filter('status ne \'completed\'')        
                .get();
    }

    public async GetUsersTaskLists (): Promise<MicrosoftGraph.TodoTaskList[]> {
        if (this.graphClient === undefined){
            throw new Error('GraphService not initialized!')
        }
        
        return await this.graphClient.api('/me/todo/lists').get();
    }

    public GetITaskItemFromToDoItem(list: MicrosoftGraph.TodoTaskList, todoItem: MicrosoftGraph.TodoTask): ITaskItem {
        let fromMail: boolean = false;
        if (list.wellknownListName === WellKnownNames.FlaggedEmails)
        {
            fromMail = true
        }

        let important: boolean = false;
        if (todoItem.importance === Importance.Important)
        {
            important = true
        }

        let dueDateString: string = ""
        if (todoItem.dueDateTime)
        {
            dueDateString = todoItem.dueDateTime.dateTime;
        }

        return  {
            fromMail: fromMail,
            listName: list.displayName,
            title: todoItem.title,
            description: todoItem.body.content,
            important: important,
            today: false,
            id: todoItem.id,
            // used for ordering
            createdDate: new Date(todoItem.createdDateTime),
            dueDate: dueDateString,
            platform: TaskPlatform.ToDo
        };
    }

    public async GetOutStandingTaskFromToDo (): Promise<ITaskItem[]> {
        let taskListItems: ITaskItem[] = [] 

        // Change to chained thens

        var taskLists = await this.GetUsersTaskLists()

        taskLists.forEach(async list => {
            var tasks = await this.GetOutstandingTodoItemsInList(list);
            
            tasks.forEach(task => {
                // Create ITaskListItem and add it to the array
                taskListItems.push(this.GetITaskItemFromToDoItem(list, task))

            })
            
        });

        return taskListItems;
    }
}


export const graphService = new GraphService();