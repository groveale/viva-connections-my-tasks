import {
  BaseImageCardView,
  IImageCardParameters,
  IExternalLinkCardAction,
  IQuickViewCardAction,
  ICardButton
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'TasksTodoAdaptiveCardExtensionStrings';
import { ITaskItem } from '../../../common/models/ITask';
import { ITasksTodoAdaptiveCardExtensionProps, ITasksTodoAdaptiveCardExtensionState, QUICK_VIEW_REGISTRY_ID } from '../TasksTodoAdaptiveCardExtension';

export class CardView extends BaseImageCardView<ITasksTodoAdaptiveCardExtensionProps, ITasksTodoAdaptiveCardExtensionState> {
  /**
   * Buttons will not be visible if card size is 'Medium' with Image Card View.
   * It will support up to two buttons for 'Large' card size.
   */
  public get cardButtons(): [ICardButton] | [ICardButton, ICardButton] | undefined {
    return [
      {
        title: strings.QuickViewButton,
        action: {
          type: 'QuickView',
          parameters: {
            view: QUICK_VIEW_REGISTRY_ID
          }
        }
      }
    ];
  }

  public get data(): IImageCardParameters {

    var allTasks = this.state.toDoTasks.concat(this.state.plannerTasks)
    var primaryString = `${allTasks.length} outstanding tasks`
    
    // if allTasks is empty, return a message
    if (allTasks.length === 0) {
      primaryString = "No outstanding tasks 😎"
    }

    
    return {
      primaryText: primaryString,
      imageUrl: `https://reckittstorage.blob.core.windows.net/viva-connections-icons/mytasksimage.svg`,
      title: this.properties.title
    };
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return {
      type: 'QuickView',
        parameters: {
          view: QUICK_VIEW_REGISTRY_ID
        }
    };
  }
}
