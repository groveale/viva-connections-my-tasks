import * as Task from "azure-devops-node-api/TaskApi";

//https://github.com/microsoft/azure-devops-node-api/blob/master/samples/task.ts

export interface IADOService {
    Init: (taskAPIClient: Task.ITaskApi, loggedinUsersUPN: string) => void;    
}

export class ADOService implements IADOService {
    
    
    //public UsersPlannerPlans: MicrosoftGraph.PlannerPlan[] = []
    private taskAPIClient: Task.ITaskApi;
    private upn: string
    private flaggedEmailListId: string
    
    public Init(taskAPIClient: Task.ITaskApi, loggedinUsersUPN: string): void {
        this.taskAPIClient = taskAPIClient
        this.upn = loggedinUsersUPN
    }

}

export const graphService = new ADOService();