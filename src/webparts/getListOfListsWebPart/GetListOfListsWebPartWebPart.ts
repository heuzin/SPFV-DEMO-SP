import {
  Environment,
  EnvironmentType,
  Version,
} from '@microsoft/sp-core-library'
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane'
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base'
import { escape } from '@microsoft/sp-lodash-subset'

import styles from './GetListOfListsWebPartWebPart.module.scss'
import * as strings from 'GetListOfListsWebPartWebPartStrings'

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http'

export interface IGetListOfListsWebPartWebPartProps {
  description: string
}

export interface ISharePointList {
  Title: string;
  Id: string;
}

export interface ISharePointLists {
  value: ISharePointList[];
}

export default class GetListOfListsWebPartWebPart extends BaseClientSideWebPart<IGetListOfListsWebPartWebPartProps> {
  private _getListOfLists(): Promise<ISharePointLists> {
    return this.context.spHttpClient
      .get(
        this.context.pageContext.web.absoluteUrl +
          `/_api/web/lists?$$filter=Hidden eq false`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json()
      })
  }

  private _getAndRenderLists(): void {
    // Local enviroment
    if (Environment.type === EnvironmentType.Local) {
    } else if (
      Environment.type === EnvironmentType.SharePoint ||
      Environment.type === EnvironmentType.ClassicSharePoint
    ) {
      this._getListOfLists().then((response) => {
        this._renderListsOfLists(response.value)
      })
    }
  }

  private _renderListsOfLists(items: ISharePointList[]): void {
    let html: string = ''

    items.forEach((item: ISharePointList) => {
      html += `
      <ul class="$${styles.list}">
        <li class="${styles.listItem}">
          <span class="ms-font-1">$${item.Title}</span>
        </li>
        <li class="$${styles.listItem}">
          <span class="ms-font-1">${item.Id}</span>
        </li>
      </ul>
      `
    })
    const listsPlaceHolder: Element =
      this.domElement.querySelector('#SPListPlaceHolder')
    listsPlaceHolder.innerHTML = html
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.getListOfListsWebPart}">
        <div class="${styles.container}">
          <div class="${styles.row}">
            <div class="${styles.column}">
              <span class="${styles.title}">Welcome to SharePoint!</span>
              <p class="${
                styles.subTitle
              }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${styles.description}">${escape(
      this.properties.description
    )}</p>
              <a href="https://aka.ms/spfx" class="${styles.button}">
                <span class="${styles.label}">Learn more</span>
              </a>
            </div>
          </div>
        </div>

        <div id='SPListPlaceHolder'</div>
      </div>`

    this._getAndRenderLists()
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0')
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    }
  }
}
