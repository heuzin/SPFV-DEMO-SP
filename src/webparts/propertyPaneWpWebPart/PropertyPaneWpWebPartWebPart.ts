import { Version } from '@microsoft/sp-core-library'
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane'
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base'
import { escape } from '@microsoft/sp-lodash-subset'

import styles from './PropertyPaneWpWebPartWebPart.module.scss'
import * as strings from 'PropertyPaneWpWebPartWebPartStrings'

export interface IPropertyPaneWpWebPartWebPartProps {
  description: string
  productName: string
  productDescription: string
  productCost: number
  quantity: number
  billAmount: number
  discount: number
  netBillAmount: number
}

export default class PropertyPaneWpWebPartWebPart extends BaseClientSideWebPart<IPropertyPaneWpWebPartWebPartProps> {
  protected onInit(): Promise<void> {
    return new Promise<void>((resolve, _reject) => {
      this.properties.productName = 'Mouse'
      ;(this.properties.productDescription = ' Mouse Description'),
        (this.properties.productCost = 0),
        (this.properties.quantity = 0),
        (this.properties.billAmount = 0),
        (this.properties.discount = 0),
        (this.properties.netBillAmount = 0)
      resolve(undefined)
    })
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.propertyPaneWpWebPart}">
        <div class="${styles.container}">
          <div class="${styles.row}">
            <div class="${styles.column}">

            <table>

            <tr>
              <td>Product Name</td>
              <td>${this.properties.productName}</td>
            </tr>

            <tr>
              <td>Description</td>
              <td>${this.properties.productDescription}</td>
            </tr>

            <tr>
              <td>Product Cost</td>
              <td>${this.properties.productCost}</td>
            </tr>

            <tr>
              <td>Product Quantity</td>
              <td>${this.properties.quantity}</td>
            </tr>

            <tr>
              <td>Bill Amount</td>
              <td>${(this.properties.billAmount =
                this.properties.productCost * this.properties.quantity)}</td>
            </tr>

            <tr>
              <td>Discount</td>
              <td>${(this.properties.discount =
                (this.properties.billAmount * 10) / 100)}</td>
            </tr>

            <tr>
              <td>Net Bill Amount</td>
              <td>${(this.properties.netBillAmount =
                this.properties.billAmount - this.properties.discount)}</td>
            </tr>

            </div>
          </div>
        </div>
      </div>`
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0')
  }

  // protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
  //   return {
  //     pages: [
  //       {
  //         header: {
  //           description: strings.PropertyPaneDescription
  //         },
  //         groups: [
  //           {
  //             groupName: strings.BasicGroupName,
  //             groupFields: [
  //               PropertyPaneTextField('description', {
  //                 label: strings.DescriptionFieldLabel
  //               })
  //             ]
  //           }
  //         ]
  //       }
  //     ]
  //   };
  // }
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: 'Product Details',
              groupFields: [
                PropertyPaneTextField('productName', {
                  label: 'Product Name',
                  multiline: false,
                  resizable: false,
                  deferredValidationTime: 5000,
                  placeholder: 'Please enter product name',
                  description: 'Name property field',
                }),

                PropertyPaneTextField('productDescription', {
                  label: 'Product Description',
                  multiline: true,
                  resizable: false,
                  deferredValidationTime: 5000,
                  placeholder: 'Please enter product description',
                  description: 'Name property field',
                }),

                PropertyPaneTextField('productCost', {
                  label: 'Product Cost',
                  multiline: false,
                  resizable: false,
                  deferredValidationTime: 5000,
                  placeholder: 'Please enter product cost',
                  description: 'Number property field',
                }),

                PropertyPaneTextField('quantity', {
                  label: 'Product Quantity',
                  multiline: false,
                  resizable: false,
                  deferredValidationTime: 5000,
                  placeholder: 'Please enter product quantity',
                  description: 'Number property field',
                }),
              ],
            },
          ],
        },
      ],
    }
  }
}
