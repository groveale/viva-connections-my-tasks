import { ISPFxAdaptiveCard, BaseAdaptiveCardView, IActionArguments, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'TasksTodoAdaptiveCardExtensionStrings';
import { ITasksTodoAdaptiveCardExtensionProps, ITasksTodoAdaptiveCardExtensionState } from '../TasksTodoAdaptiveCardExtension';
import { graphService } from '../../../common/services/MSGraphService';
import { TaskPlatform } from '../../../common/models/ITask';

export interface IDetailedQuickViewData {
  task: any;
  allTasks: any[];
  strings: ITasksTodoAdaptiveCardExtensionStrings;
}

export class DetailedQuickView extends BaseAdaptiveCardView<
ITasksTodoAdaptiveCardExtensionProps,
    ITasksTodoAdaptiveCardExtensionState,
    IDetailedQuickViewData
> {
  public get data(): IDetailedQuickViewData {
    var allTasks = this.state.toDoTasks.concat(this.state.plannerTasks, this.state.adoTasks);
    const tasks = allTasks.filter((task: any) => {
        return task.id === this.state.currentTaskKey;
      });
    return {
      task: tasks[0],
      allTasks: allTasks,
      strings: strings,
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/DetailedViewTemplate.json');
  }

  public async onAction(action: IActionArguments): Promise<void> {
    if ((<ISubmitActionArguments>action).type === 'Submit') {
      const submitAction = <ISubmitActionArguments>action;
      const { id, taskKey } = submitAction.data;
      if (id === 'closeTask') {
        // We actually need to mark as complete in todo
        graphService.MarkTaskAsDone(this.data.task);
        this.setState({ 
            toDoTasks: this.state.toDoTasks.filter((item: any) => item.id !== taskKey),
            plannerTasks: this.state.plannerTasks.filter((item: any) => item.id !== taskKey),
        });
        this.quickViewNavigator.pop();
        }
    }
    }
}