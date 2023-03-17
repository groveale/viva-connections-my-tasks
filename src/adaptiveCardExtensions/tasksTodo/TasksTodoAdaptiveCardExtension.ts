import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { TasksTodoPropertyPane } from './TasksTodoPropertyPane';
import { graphService } from '../../common/services/MSGraphService';
import { ITaskItem } from '../../common/models/ITask';

export interface ITasksTodoAdaptiveCardExtensionProps {
  title: string;
}

export interface ITasksTodoAdaptiveCardExtensionState {
  toDoTasks: ITaskItem[]
}

const CARD_VIEW_REGISTRY_ID: string = 'TasksTodo_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'TasksTodo_QUICK_VIEW';

export default class TasksTodoAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  ITasksTodoAdaptiveCardExtensionProps,
  ITasksTodoAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: TasksTodoPropertyPane | undefined;

  public async onInit(): Promise<void> {
    this.state = {
        toDoTasks: []
     };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

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
    setTimeout(async () => {
      try {
        const todoTasks: ITaskItem[] = await graphService.GetOutStandingTaskFromToDo();
        this.setState({
          toDoTasks: todoTasks
        });
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