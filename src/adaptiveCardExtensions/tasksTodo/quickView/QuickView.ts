import { ISPFxAdaptiveCard, BaseAdaptiveCardView, IActionArguments, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'TasksTodoAdaptiveCardExtensionStrings';
import { COMING_SOON_VIEW_REGISTRY_ID, DETAILED_VIEW_REGISTRY_ID, ITasksTodoAdaptiveCardExtensionProps, ITasksTodoAdaptiveCardExtensionState } from '../TasksTodoAdaptiveCardExtension';

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
    //var allTasks = this.state.toDoTasks.concat(this.state.plannerTasks, this.state.adoTasks);
    var allTasks = this.state.toDoTasks.concat(this.state.plannerTasks);
    if (allTasks.length > 1) {
      numberOfTasks = `${allTasks.length.toString()} ${strings.CardViewTextPlural}`;
    } else {
      numberOfTasks = `${allTasks.length.toString()} ${strings.CardViewTextSingular}`;
    }



    return {
      numberOfTasks: numberOfTasks,
      tasks: allTasks.sort((a, b) => b.createdDate.getTime() - a.createdDate.getTime()),
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
        if (taskKey === "devops-coming-soon") {
          this.quickViewNavigator.push(COMING_SOON_VIEW_REGISTRY_ID);
        } else {
          this.quickViewNavigator.push(DETAILED_VIEW_REGISTRY_ID);
        }
        
      }
    }
  }
}