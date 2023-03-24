// import * as azdev from "azure-devops-node-api";
// import * as msal from '@azure/msal-node';
// import { WorkItem } from "azure-devops-node-api/interfaces/WorkItemTrackingInterfaces";
// import { ITaskItem, PlatformLogo, TaskPlatform } from "../models/ITask";

// //https://github.com/microsoft/azure-devops-node-api/blob/master/samples/task.ts

// export interface IADOService {
//     Init: (orgName: string, loggedinUsersUPN: string) => void;
//     GetWorkItems: () => Promise<WorkItem[]>;    
// }

// export class AzureDevOpsService implements IADOService {

//     private _connection: azdev.WebApi;
//     private orgName: string = "groveale"
//     private loggedinUsersUPN: string
    
//     public Init(orgName: string, loggedinUsersUPN: string): void {
//         this.orgName = orgName,
//         this.loggedinUsersUPN = loggedinUsersUPN
//     }


//     public async getConnectionPersonal(): Promise<azdev.WebApi> {
//       if (!this._connection) {
//         const orgUrl = `https://dev.azure.com/${this.orgName}`;
//         const token = process.env.AZURE_DEVOPS_EXT_PAT;
//         const authHandler = azdev.getPersonalAccessTokenHandler(token);
//         this._connection = new azdev.WebApi(orgUrl, authHandler);
//       }
//       return this._connection;
//     }

//     public async getConnectionApp(): Promise<azdev.WebApi> {
//         if (!this._connection) {
//             const clientId = 'ff9d3c0d-f485-4eaf-9849-fcb78ad5df8a';
//             const clientSecret = '6Q28Q~hFB2ZYeEaAgUFhPIesrVykIIo3rRDfCcJJ';
//             const tenantId = '75e67881-b174-484b-9d30-c581c7ebc177';
//             const orgUrl = `https://dev.azure.com/${this.orgName}`;

//             // Create a new MSAL ConfidentialClientApplication
//             const pca = new msal.ConfidentialClientApplication({
//                 auth: {
//                 clientId: clientId,
//                 authority: `https://login.microsoftonline.com/${tenantId}`,
//                 clientSecret: clientSecret
//                 }
//             });

//             // Get an access token for the Azure DevOps API using the MSAL application
//             const tokenResponse = await pca.acquireTokenByClientCredential({
//                 scopes: [`${orgUrl}/.default`]
//             });

//             // Create a connection to Azure DevOps using the access token
//             const authHandler = azdev.getBearerHandler(tokenResponse.accessToken);
//             this._connection = new azdev.WebApi(orgUrl, authHandler);
//         }
          
//         return this._connection;
//     }
  
//     public async GetWorkItems(): Promise<WorkItem[]> {
      
//         // Get a connection to Azure DevOps
//         const connection = await this.getConnectionApp();

//         // Get a reference to the work item tracking API
//         const witApi = await connection.getWorkItemTrackingApi();

//         // Define the query to retrieve work items assigned to the user email
//         // Can remove the fields as we get them later
//         const query = {
//             query: `Select [System.Id], [System.Title], [System.State], [System.WorkItemType] 
//                     From WorkItems 
//                     Where [System.AssignedToEmail] = @Me`,

//             parameters: [{ name: "Me", value: this.loggedinUsersUPN }],
//         };

//         // Run the query to retrieve the work items
//         const result = await witApi.queryByWiql(query);

//         const workItemIds = result.workItems.map((wi) => wi.id);
//         const fields = ["System.Id", "System.Title", "System.State", "System.WorkItemType" ];
//         const workItems = await witApi.getWorkItems(workItemIds, fields);

//         return workItems;
//     }

//     public GetITaskItemFromADOWorkItem(adoWorkItem: WorkItem ): ITaskItem {

//         let dueDateString: string = ""
//         let overDueDays: string = ""
//         let planTitle: string = ""

//         return  {
//             listName: planTitle,
//             title: adoWorkItem.fields['System.Title'],
//             //description: plannerItem.details.description,
//             today: false,
//             id: adoWorkItem.id.toString(),
//             dueDate: dueDateString,
//             platform: TaskPlatform.AzureDevOps,
//             logoUrl: PlatformLogo.AzureDevOps,
//             overDueDays: overDueDays
//         };
//     }
// }

// export const adoService = new AzureDevOpsService();