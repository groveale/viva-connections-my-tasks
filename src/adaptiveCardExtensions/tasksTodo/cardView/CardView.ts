import {
  BaseImageCardView,
  IImageCardParameters,
  IExternalLinkCardAction,
  IQuickViewCardAction,
  ICardButton
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'TasksTodoAdaptiveCardExtensionStrings';
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
    return {
      primaryText: `${this.state.toDoTasks.length} outstanding tasks`,
      imageUrl: require('../assets/MicrosoftLogo.png'),
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
