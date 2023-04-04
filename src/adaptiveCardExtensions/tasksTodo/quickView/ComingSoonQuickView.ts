import { ISPFxAdaptiveCard, BaseAdaptiveCardView, IActionArguments, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'TasksTodoAdaptiveCardExtensionStrings';
import { ITasksTodoAdaptiveCardExtensionProps, ITasksTodoAdaptiveCardExtensionState } from '../TasksTodoAdaptiveCardExtension';

export interface IComingSoonQuickViewData {
  task: any;
  strings: ITasksTodoAdaptiveCardExtensionStrings;
}

export class ComingSoonQuickView extends BaseAdaptiveCardView<
ITasksTodoAdaptiveCardExtensionProps,
    ITasksTodoAdaptiveCardExtensionState,
    IComingSoonQuickViewData
> {
  public get data(): IComingSoonQuickViewData {
    const tasks = this.state.toDoTasks.filter((task: any) => {
        return task.id === this.state.currentTaskKey;
      });
    return {
      task: tasks[0],
      strings: strings,
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/ComingSoonTemplate.json');
  }
}