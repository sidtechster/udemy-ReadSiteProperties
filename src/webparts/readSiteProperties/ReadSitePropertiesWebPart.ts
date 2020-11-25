import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneCheckbox,
  PropertyPaneLink
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ReadSitePropertiesWebPart.module.scss';
import * as strings from 'ReadSitePropertiesWebPartStrings';

import {Environment, EnvironmentType} from '@microsoft/sp-core-library';

import {SPHttpClient, SPHttpClientResponse} from '@microsoft/sp-http';

export interface IReadSitePropertiesWebPartProps {
  description: string;
  envTitle: string;

  productName: string;
  productDescription: string;
  productCost: number;
  quantity: number;
  billAmount: number;
  discount: number;
  netBillAmount: number;

  isCertified: boolean;
  rating: number;
  processorType: string;
  invoiceFileType: string;
  newProcessorType: string;
  discountCoupon: boolean;

}

export interface ISharePointList {
  Title: string;
  Id: string;
}

export interface ISharePointLists {
  value: ISharePointList[];
}

export default class ReadSitePropertiesWebPart extends BaseClientSideWebPart<IReadSitePropertiesWebPartProps> {

  protected onInit(): Promise<void> {
    return new Promise<void>((resolve, _reject) => {
      this.properties.productName = "Macbook Air";
      this.properties.productDescription = "M1 chip with 256GB storage";
      this.properties.quantity = 1;
      this.properties.productCost = 92900;

      resolve(undefined);

    });
  }

  protected get disableReactivePropertyChanges(): boolean {
    return false;
  }

  private _getListOfLists(): Promise<ISharePointLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=Hidden eq false`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private _getAndRenderLists(): void {
    if (Environment.type == EnvironmentType.Local) {
      
    }
    else if (Environment.type == EnvironmentType.ClassicSharePoint
    || Environment.type == EnvironmentType.SharePoint) {
      this._getListOfLists()
        .then((response) => {
          this._renderListOfLists(response.value);
        });
    }
  }

  private _renderListOfLists(items: ISharePointList[]): void {
    let html = '';

    items.forEach((item: ISharePointList) => {
      html += `
      <ul class="${styles.list}">
        <li class="${styles.listItem}">
          <span class="ms-font-l">${item.Title}</span>
        </li>
        <li class="${styles.listItem}">
          <span class="ms-font-l">${item.Id}</span>
        </li>
      </ul>`;
    });

    const listPlaceHolder: Element = this.domElement.querySelector('#SPListPlaceHolder');
    listPlaceHolder.innerHTML = html;
  } 

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.readSiteProperties }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              
            <table>

            <tr>
              <td>Product Name</td>
              <td>${this.properties.productName}</td>
            </tr>

            <tr>
              <td>Product Description</td>
              <td>${this.properties.productDescription}</td>
            </tr>

            <tr>
              <td>Cost</td>
              <td>${this.properties.productCost}</td>
            </tr>

            <tr>
              <td>Quantity</td>
              <td>${this.properties.quantity}</td>
            </tr>

            <tr>
              <td>Bill Amount</td>
              <td>${this.properties.billAmount = this.properties.productCost * this.properties.quantity}</td>
            </tr>

            <tr>
              <td>Discount</td>
              <td>${this.properties.discount = this.properties.billAmount * 0.1}</td>
            </tr>

            <tr>
              <td>Net Bill Amount</td>
              <td>${this.properties.netBillAmount = this.properties.billAmount - this.properties.discount}</td>
            </tr>

            <tr>
              <td>Is certified?</td>
              <td>${this.properties.isCertified}</td>
            </tr>

            <tr>
              <td>Rating</td>
              <td>${this.properties.rating}</td>
            </tr>

            <tr>
              <td>Processor Type</td>
              <td>${this.properties.processorType}</td>
            </tr>

            <tr>
              <td>Invoice Type</td>
              <td>${this.properties.invoiceFileType}</td>
            </tr>

            <tr>
              <td>New Processor Type</td>
              <td>${this.properties.newProcessorType}</td>
            </tr>

            <tr>
              <td>Do you have a discount coupon</td>
              <td>${this.properties.discountCoupon}</td>
            </tr>

            </table>
              
            </div>
          </div>
          <div id="SPListPlaceHolder"></div>
        </div>
      </div>`;

      //this._getAndRenderLists();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
              groupName: "Product Details",
              groupFields: [
                PropertyPaneTextField('productName', {
                  label: "Product Name",
                  multiline: false,
                  resizable: false,
                  deferredValidationTime: 5000,
                  placeholder: "Please enter product name","description": "Name property field"
                }),
                PropertyPaneTextField('productDescription', {
                  label: "Product Description",
                  multiline: true,
                  resizable: false,
                  deferredValidationTime: 5000,
                  placeholder: "Please enter product description","description": "Name property field"
                }),
                PropertyPaneTextField('productCost', {
                  label: "Product Cost",
                  multiline: false,
                  resizable: false,
                  deferredValidationTime: 5000,
                  placeholder: "Please enter product cost","description": "Name property field"
                }),
                PropertyPaneTextField('quantity', {
                  label: "Product Quantity",
                  multiline: false,
                  resizable: false,
                  deferredValidationTime: 5000,
                  placeholder: "Please enter product quantity","description": "Name property field"
                }),
                PropertyPaneToggle('isCertified', {
                  key: 'isCertified',
                  label: 'Is it certified?',
                  onText: 'ECC certified',
                  offText: 'Not certified'
                }),
                PropertyPaneSlider('rating', {
                  label: 'Select your rating',
                  min: 1,
                  max: 5,
                  step: 1,
                  showValue: true,
                  value: 1
                }),
                PropertyPaneChoiceGroup('processorType', {
                  label: 'Processor',
                  options: [
                    {key: 'i5', text: 'Intel i5'},
                    {key: 'i7', text: 'Intel i7', checked: true},
                    {key: 'Ryzen3800x', text: 'AMD Ryzen 3800X'},
                  ]
                }),
                PropertyPaneChoiceGroup('invoiceFileType', {
                  label: 'Select invoice file type:',
                  options: [
                    {
                      key: 'MSWord', text: 'MSWord',
                      imageSrc: 'https://www.flaticon.com/svg/static/icons/svg/888/888883.svg',
                      imageSize: { width: 32, height: 32 },
                      selectedImageSrc: 'https://www.flaticon.com/svg/static/icons/svg/888/888883.svg'
                    },
                    {
                      key: 'MSExcel', text: 'MSExcel',
                      imageSrc: 'https://www.flaticon.com/svg/static/icons/svg/732/732220.svg',
                      imageSize: { width: 32, height: 32 },
                      selectedImageSrc: 'https://www.flaticon.com/svg/static/icons/svg/732/732220.svg'
                    },
                    {
                      key: 'PDF', text: 'PDF',
                      imageSrc: 'https://www.flaticon.com/svg/static/icons/svg/337/337946.svg',
                      imageSize: { width: 32, height: 32 },
                      selectedImageSrc: 'https://www.flaticon.com/svg/static/icons/svg/337/337946.svg'
                    },
                  ]
                }),
                PropertyPaneDropdown('newProcessorType', {
                  label: 'New Processor Type',
                  options: [
                    {key: 'i5', text: 'Intel i5'},
                    {key: 'i7', text: 'Intel i7'},
                    {key: 'Ryzen3800x', text: 'AMD Ryzen 3800X'},
                  ],
                  selectedKey: 'i7'
                }),
                PropertyPaneCheckbox('discountCoupon', {
                  text: 'Do you have a discount coupon?',
                  checked: false,
                  disabled: false
                }),
                PropertyPaneLink('', {
                  href: 'https://www.amazon.in',
                  text: 'Buy processors from genuine seller',
                  target: '_blank',
                  popupWindowProps: {
                    height: 500,
                    width: 500,
                    positionWindowPosition: 2,
                    title: 'Amazon'
                  }
                })
              ]
            }
          ]
        }
      ]
    }
  }

}
