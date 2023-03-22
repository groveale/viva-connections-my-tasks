import { ISPFxAdaptiveCard, BaseAdaptiveCardView, IActionArguments, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'TasksTodoAdaptiveCardExtensionStrings';
import { ITasksTodoAdaptiveCardExtensionProps, ITasksTodoAdaptiveCardExtensionState } from '../TasksTodoAdaptiveCardExtension';

export interface IDetailedQuickViewData {
  task: any;
  strings: ITasksTodoAdaptiveCardExtensionStrings;
}

export class DetailedQuickView extends BaseAdaptiveCardView<
ITasksTodoAdaptiveCardExtensionProps,
    ITasksTodoAdaptiveCardExtensionState,
    IDetailedQuickViewData
> {
  public get data(): IDetailedQuickViewData {
    const tasks = this.state.toDoTasks.filter((task: any) => {
        return task.id === this.state.currentTaskKey;
      });
    return {
      task: tasks[0],
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
        this.setState({ 
            toDoTasks: this.state.toDoTasks.filter((item: any) => item.id !== taskKey)
        });
        this.quickViewNavigator.pop();
        }
    }
    }
}