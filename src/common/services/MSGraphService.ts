import { MSGraphClientV3 } from '@microsoft/sp-http'
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { ITaskItem, PlatformLogo, TaskPlatform } from '../models/ITask';

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
    
    GetITaskItemFromToDoItem: (list: MicrosoftGraph.TodoTaskList, todoItem: MicrosoftGraph.TodoTask) => ITaskItem

    GetUsersPlannerTasks: () => Promise<MicrosoftGraph.PlannerTask[]>;
    GetITaskItemFromPlannerItem: (plannerItem: MicrosoftGraph.PlannerTask, plannerPlan: MicrosoftGraph.PlannerPlan[]) => ITaskItem
    MarkTaskAsDone: (task: ITaskItem) => Promise<void>;
}

export class GraphService implements IGraphService {
    
    
    //public UsersPlannerPlans: MicrosoftGraph.PlannerPlan[] = []
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
            .get()
            .then((response) => {
                // Need to return the actual array of tasks
                return response.value;
            })
    }

    public async GetUsersTaskLists (): Promise<MicrosoftGraph.TodoTaskList[]> {
        if (this.graphClient === undefined){
            throw new Error('GraphService not initialized!')
        }
        
        return await this.graphClient.api('/me/todo/lists').get()
            .then((response) => {
                // Need to return the actual array of task lists
                return response.value;
            })
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
        let overDueDays: string = ""
        if (todoItem.dueDateTime)
        {
            dueDateString = todoItem.dueDateTime.dateTime;
            if (new Date(todoItem.dueDateTime.dateTime).getMilliseconds() < Date.now()) {
                // Overdue
                overDueDays = Math.round((Date.now() - new Date(todoItem.dueDateTime.dateTime).getMilliseconds()) / (1000 * 60 * 60 * 24)).toString();
            }
        }

        return  {
            fromMail: fromMail,
            listName: list.displayName,
            listId: list.id,
            title: todoItem.title,
            description: todoItem.body.content,
            important: important,
            today: false,
            id: todoItem.id,
            // used for ordering
            createdDate: new Date(todoItem.createdDateTime),
            // because adaptive cards does not support ISO 8601 format
            createdDateString: (new Date(todoItem.createdDateTime)).toISOString().substr(0, 19) + "Z",
            dueDate: dueDateString,
            platform: TaskPlatform.ToDo,
            logoUrl: PlatformLogo.ToDo,
            overDueDays: overDueDays,
            source: "ToDo",
            deepLinkUrl: this.GetDeepLinkUrlForToDoTask(todoItem.id)

        };
    }


    private GetDeepLinkUrlForToDoTask(taskId: string): string {
        
        const teamsAppId = 'com.microsoft.teamspace.tab.planner'; // The app ID for Tasks by Planner and To Do
        const deepLink = `msteams://teams.microsoft.com/l/entity/${teamsAppId}/tasks?webUrl=https%3A%2F%2Fto-do.office.com`;

        return deepLink;
    }


    public async MarkTaskAsDone(task: ITaskItem): Promise<void> {
        if (this.graphClient === undefined){
            throw new Error('GraphService not initialized!')
        }


        switch (task.platform) {
            case TaskPlatform.ToDo:
                await this.graphClient.api(`/me/todo/lists/${task.listId}/tasks/${task.id}`)
                    .patch({
                        status: "completed"
                    })
                break;
            case TaskPlatform.Planner:
                // need more permissions to do this
                // await this.graphClient.api(`/planner/tasks/${task.id}`)
                //     .patch({
                //         percentComplete: 100
                //     })
                break;
            default:
                // do nothing
        }
        
    }

    public async GetUsersPlannerTasks(): Promise<MicrosoftGraph.PlannerTask[]> {
        if (this.graphClient === undefined){
            throw new Error('GraphService not initialized!')
        }

        // return tasks that are not completed in the list
        return await this.graphClient.api(`/me/planner/tasks`)    
            .get()
            .then((response) => {
                // Need to return the actual array of planner tasks
                return response.value;
            })
    }

    public async GetUsersPlanerPlans(): Promise<MicrosoftGraph.PlannerPlan[]> {
        if (this.graphClient === undefined){
            throw new Error('GraphService not initialized!')
        }

        return await this.graphClient.api('/me/planner/plans')
            // filter does not work with this call :S
            //.filter(`id eq \'${task.planId}\'`)
            .get()
            .then((response) => {
                return response.value;
            })
    
    }

    public GetITaskItemFromPlannerItem(plannerItem: MicrosoftGraph.PlannerTask, plannerPlan: MicrosoftGraph.PlannerPlan[]): ITaskItem {


        let dueDateString: string = ""
        let overDueDays: string = ""
        let planTitle: string = ""
        if (plannerItem.dueDateTime)
        {
            dueDateString = plannerItem.dueDateTime;
            if (new Date(plannerItem.dueDateTime).getMilliseconds() < Date.now()) {
                // Overdue
                overDueDays = Math.round((Date.now() - new Date(plannerItem.dueDateTime).getMilliseconds()) / (1000 * 60 * 60 * 24)).toString();
            }
        }

        if (plannerPlan.length > 0) {
            planTitle = plannerPlan[0].title
        }

        return  {
            listName: planTitle,
            title: plannerItem.title,
            description: "",
            today: false,
            id: plannerItem.id,
            // used for ordering
            createdDate: new Date(plannerItem.createdDateTime),
            // because adaptive cards does not support ISO 8601 format
            createdDateString: (new Date(plannerItem.createdDateTime)).toISOString().substr(0, 19) + "Z",
            dueDate: dueDateString,
            platform: TaskPlatform.Planner,
            logoUrl: PlatformLogo.Planner,
            overDueDays: overDueDays,
            source: "Planner",
            percentComplete: plannerItem.percentComplete,
            deepLinkUrl: this.GetDeepLinkUrlForToDoTask(plannerItem.id)
        };
    }

    
}


export const graphService = new GraphService();