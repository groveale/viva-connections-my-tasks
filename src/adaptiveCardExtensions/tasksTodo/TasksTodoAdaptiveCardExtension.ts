import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { TasksTodoPropertyPane } from './TasksTodoPropertyPane';
import { graphService } from '../../common/services/MSGraphService';
import { ITaskItem } from '../../common/models/ITask';
import { DetailedQuickView } from './quickView/DetailedQuickView';
import { PlannerPlan } from '@microsoft/microsoft-graph-types';

export interface ITasksTodoAdaptiveCardExtensionProps {
  title: string;
}

export interface ITasksTodoAdaptiveCardExtensionState {
  toDoTasks: ITaskItem[]
  plannerTasks: ITaskItem[]
  currentTaskKey: string;
}

const CARD_VIEW_REGISTRY_ID: string = 'TasksTodo_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'TasksTodo_QUICK_VIEW';
export const DETAILED_VIEW_REGISTRY_ID: string = 'TasksTodo_DETAILED_VIEW'


export default class TasksTodoAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  ITasksTodoAdaptiveCardExtensionProps,
  ITasksTodoAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: TasksTodoPropertyPane | undefined;

  public async onInit(): Promise<void> {
    this.state = {
        toDoTasks: [],
        plannerTasks: [],
        currentTaskKey: ""
     };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());
    this.quickViewNavigator.register(DETAILED_VIEW_REGISTRY_ID, () => new DetailedQuickView());


    const graphClient = await this.context.msGraphClientFactory.getClient("3");

    // Graph service for a clean design
    graphService.Init(graphClient, this.context.pageContext.user.email);

    return this.GetTaskData();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'TasksTodo-property-pane'*/
      './TasksTodoPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.TasksTodoPropertyPane();
        }
      );
  }

  private GetTaskData(): Promise<void> {
    
    // dedicated requests to speed up the load of the Adaptive Card (async)

    // Get ToDo Tasks
    setTimeout(async () => {
      try {

        const todoTasks: ITaskItem[] = []
        let index = 0;
        var taskLists = await graphService.GetUsersTaskLists()

        taskLists.forEach(async list => {
            var tasks = await graphService.GetOutstandingTodoItemsInList(list);
            
            tasks.forEach(task => {
                // Create ITaskListItem and add it to the array
                todoTasks.push(graphService.GetITaskItemFromToDoItem(list, task))
                
                this.setState({
                  toDoTasks: todoTasks
                });
            })
            
        });
      } catch (error) {
        console.error(error);
      }
    }, 500)


    // Get Planner Tasks
    setTimeout(async () => {
      try {

        const plannerTasks: ITaskItem[] = []
        let index = 0;
        var planerPlans = await graphService.GetUsersPlanerPlans()
        var plansFromPlanner = await graphService.GetUsersPlannerTasks()

        plansFromPlanner.forEach(async plan => {
            // Create ITaskListItem and add it to the array

            var plannerPlan: PlannerPlan[] = planerPlans.filter((item: PlannerPlan) => item.id === plan.planId)

            plannerTasks.push(graphService.GetITaskItemFromPlannerItem(plan, plannerPlan))
            
            this.setState({
              plannerTasks: plannerTasks
            });
          })
      } catch (error) {
        console.error(error);
      }
    }, 500)

    return Promise.resolve();
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration();
  }
}
