import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer, PlaceholderName, PlaceholderContent } from '@microsoft/sp-application-base';
import { SPPermission } from '@microsoft/sp-page-context';
import * as strings from 'ApplicationApplicationCustomizerStrings';

const LOG_SOURCE: string = 'ApplicationApplicationCustomizer';

export interface IApplicationApplicationCustomizerProperties {
  testMessage: string;
}

export default class ApplicationApplicationCustomizer
  extends BaseApplicationCustomizer<IApplicationApplicationCustomizerProperties> {
  private _topPlaceholder: PlaceholderContent | undefined;

  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);

    const canEdit: boolean = this.context.pageContext.web.permissions.hasPermission(SPPermission.manageWeb) ||
      this.context.pageContext.legacyPageContext.isSiteAdmin;
    console.log("User has 'manageWeb' permission?", canEdit);

    const checkExist = setInterval(() => {
      const element = document.getElementById("sp-appBar");
      if (element) {
        if (canEdit) {
          console.log("User has edit permissions — App Bar stays.");
        } else {
          element.remove();
          console.log("GT user only has view rights — App Bar removed.");
        }
        clearInterval(checkExist);
      }
    }, 100);

    return Promise.resolve();
  }

  private _renderPlaceHolders(): void {
    console.log('Available placeholders: ', this.context.placeholderProvider.placeholderNames);

    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._onDispose }
      );

      if (!this._topPlaceholder) {
        console.error('The expected placeholder was not found.');
        return;
      }

      if (this.properties) {
        // inject the CSS
        if (this._topPlaceholder.domElement) {
          this._topPlaceholder.domElement.innerHTML = `
            <style>
              .j_b_4ade22aa {
                display: none !important;
              }
            </style>`;
        }
      }
    }
  }

  private _onDispose(): void {
    console.log('Disposed social bar.');
  }
}