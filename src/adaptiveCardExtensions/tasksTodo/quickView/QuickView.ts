import { ISPFxAdaptiveCard, BaseAdaptiveCardView, IActionArguments, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'TasksTodoAdaptiveCardExtensionStrings';
import { DETAILED_VIEW_REGISTRY_ID, ITasksTodoAdaptiveCardExtensionProps, ITasksTodoAdaptiveCardExtensionState } from '../TasksTodoAdaptiveCardExtension';

export interface IQuickViewData {
  numberOfTasks: string;
  tasks: object[];
  strings: ITasksTodoAdaptiveCardExtensionStrings;
}

export class QuickView extends BaseAdaptiveCardView<
  ITasksTodoAdaptiveCardExtensionProps,
  ITasksTodoAdaptiveCardExtensionState,
  IQuickViewData
> {
  public get data(): IQuickViewData {
    let numberOfTasks: string = strings.CardViewZero;
    if (this.state.toDoTasks.length > 1) {
      numberOfTasks = `${this.state.toDoTasks.length.toString()} ${strings.CardViewTextPlural}`;
    } else {
      numberOfTasks = `${this.state.toDoTasks.length.toString()} ${strings.CardViewTextSingular}`;
    }
    return {
      numberOfTasks: numberOfTasks,
      tasks: this.state.toDoTasks,
      strings: strings,
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/MyTasksList.json');
  }

  public async onAction(action: IActionArguments): Promise<void> {
    if ((<ISubmitActionArguments>action).type === 'Submit') {
      const submitAction = <ISubmitActionArguments>action;
      const { id, taskKey } = submitAction.data;
      if (id === 'selectTask') {
        this.setState({ currentTaskKey: taskKey });
        this.quickViewNavigator.push(DETAILED_VIEW_REGISTRY_ID);
      }
    }
  }
}