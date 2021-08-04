import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneCheckbox,
  PropertyPaneLink,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './PropertyPaneWpWebPartWebPart.module.scss';
import * as strings from 'PropertyPaneWpWebPartWebPartStrings';

export interface IPropertyPaneWpWebPartWebPartProps {
  description: string;
  productName: string;
  productDescription: string;
  productCost: number;
  quantity: number;
  billAmount: number;
  discount: number;
  netBillAmount: number;

  currentTime: Date;
  IsCertified: boolean;
  Rating: number;
  ProcessorType: string;
  InvoiceFileType: string;
  NewProcessorType: string;
  DiscountCouppon: boolean;
}

export default class PropertyPaneWpWebPartWebPart extends BaseClientSideWebPart<IPropertyPaneWpWebPartWebPartProps> {
  protected onInit(): Promise<void> {
    return new Promise<void>((resolve, _reject) => {
      this.properties.productName = 'Mouse';
      (this.properties.productDescription = ' Mouse Description'),
        (this.properties.productCost = 0),
        (this.properties.quantity = 0),
        (this.properties.billAmount = 0),
        (this.properties.discount = 0),
        (this.properties.netBillAmount = 0);
      this.properties.IsCertified = false;
      this.properties.currentTime = new Date();
      this.properties.Rating = 1;
      this.properties.ProcessorType = 'I7';
      this.properties.InvoiceFileType = 'MSPowerPoint';
      this.properties.NewProcessorType = 'I7';
      this.properties.DiscountCouppon = false;
      resolve(undefined);
    });
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.propertyPaneWpWebPart}">
        <div class="${styles.container}">
          <div class="${styles.row}">
            <div class="${styles.column}">

            <table>

            <tr>
              <td>Current Time</td>
              <td>${this.properties.currentTime}</td>
            </tr>

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

            <tr>
                <td>Is Certified</td>
                <td>${this.properties.IsCertified}</td>
            </tr>

            <tr>
              <td>Rating</td>
              <td>${this.properties.Rating}</td>
            </tr>

            <tr>
              <td>Processor Type</td>
              <td>${this.properties.ProcessorType}</td>
            </tr>

            <tr>
              <td>Document Type</td>
              <td>${this.properties.InvoiceFileType}</td>
            </tr>

            <tr>
              <td>New Processor Type</td>
              <td>${this.properties.NewProcessorType}</td>
            </tr>

            <tr>
              <td>Discount Coupon</td>
              <td>${this.properties.DiscountCouppon}</td>
            </tr>

            </table>

            </div>
          </div>
        </div>
      </div>`;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Inventory Web Part',
          },
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

                PropertyPaneToggle('IsCertified', {
                  key: 'IsCertfied',
                  label: 'Is it Certified?',
                  onText: 'ISI Certified!',
                  offText: 'Not an ISI Certified Product',
                }),

                PropertyPaneSlider('Rating', {
                  label: 'Rating',
                  min: 1,
                  max: 10,
                  step: 1,
                  showValue: true,
                  value: 1,
                }),

                PropertyPaneChoiceGroup('ProcessorType', {
                  label: 'Choices',
                  options: [
                    { key: 'I5', text: 'Intel I5' },
                    { key: 'I7', text: 'Intel I7', checked: true },
                    { key: 'I9', text: 'Intel I9' },
                  ],
                }),

                PropertyPaneChoiceGroup('InvoiceFileType', {
                  label: 'Select Invoice File type',
                  options: [
                    {
                      key: 'MSWord',
                      text: 'MSWord',
                      imageSrc:
                        'https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/png/docx_32x1.png',
                      imageSize: { width: 32, height: 32 },
                      selectedImageSrc:
                        'https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/png/docx_32x1.png',
                    },
                    {
                      key: 'MSExel',
                      text: 'MSExel',
                      imageSrc:
                        'https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/png/xlsx_32x1.png',
                      imageSize: { width: 32, height: 32 },
                      selectedImageSrc:
                        'https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/png/xlsx_32x1.png',
                    },
                    {
                      key: 'MSPowerPoint',
                      text: 'MSPowerPoint',
                      imageSrc:
                        'https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/png/pptx_32x1.png',
                      imageSize: { width: 32, height: 32 },
                      selectedImageSrc:
                        'https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/png/pptx_32x1.png',
                    },
                  ],
                }),

                PropertyPaneDropdown('NewProcessorType', {
                  label: 'New Processor Type',
                  options: [
                    { key: 'I5', text: 'Intel I5' },
                    { key: 'I7', text: 'Intel I7' },
                    { key: 'I9', text: 'Intel I9' },
                  ],
                }),

                PropertyPaneCheckbox('DiscountCouppon', {
                  text: 'Do you have a discount coupon?',
                  checked: false,
                  disabled: false,
                }),

                PropertyPaneLink('', {
                  href: 'https:/www.amazon.in',
                  text: 'Buy Intel Processor from the best seller',
                  target: '_blank',
                  popupWindowProps: {
                    height: 500,
                    width: 500,
                    positionWindowPosition: 2,
                    title: 'Amazons',
                  },
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
