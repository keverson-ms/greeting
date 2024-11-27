import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './JsGreetingWebPart.module.scss';
import * as strings from 'JsGreetingWebPartStrings';

export interface IJsGreetingWebPartProps {
  greeting: string;
  alignment: string;
  border: string;
}

export default class JsGreetingWebPart extends BaseClientSideWebPart<IJsGreetingWebPartProps> {

  public render(): void {
    const photo = `${escape(this.context.pageContext.site.absoluteUrl)}/_layouts/15/userphoto.aspx?size=L&accountname=${escape(this.context.pageContext.user.email)}`;

    this.domElement.innerHTML = `
    <div class="${styles.flex} ${styles.alignItemsCenter}">
      <div class="${this.properties.alignment} ${styles.w15}">
        <img class="${styles.perfil}" src="${photo}" style="border-radius: ${this.properties.border}%" />
      </div>
      <div class="${styles.jsGreeting} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''} ${this.properties.alignment} ${styles.w85}">
          <h2 class="${styles.m0}">${escape(this.properties.greeting)} <br />${escape(this.context.pageContext.user.displayName.split(' -').shift())}.</h2>
      </div>
    </div>`;
  }
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private validateBorderValue(value: string): string {
    const regex = /^[0-9]+$/;
    const numValue = parseInt(value, 10);

    if (!regex.test(value.toString())) {
      return 'O valor deve ser um número inteiro de 0 a 50.';
    }

    if (!regex.test(value.toString()) && numValue < 0 || numValue > 50) {
      return 'O valor deve estar entre 0 e 50.';
    }

    return '';
  }

  public onInit(): Promise<void> {
    if (!this.properties.greeting || this.properties.greeting.trim() === '') {
      this.properties.greeting = 'Olá, ';
    }

    if (!this.properties.border || this.properties.border.trim() === '') {
      this.properties.border = '50';
    }

    return super.onInit();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneGreeting
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('greeting', {
                  label: strings.GreetingFieldLabel,
                  placeholder: 'Digite uma saudação ou deixe em branco para usar "Olá, "',
                  description: 'Digite uma saudação ou deixe em branco para usar "Olá, "',
                }),
                PropertyPaneTextField('border', {
                  label: `${strings.BorderFieldLabel}`,
                  placeholder: 'Insira um valor de 0 a 50',
                  description: 'Insira um valor de 0 a 50',
                  onGetErrorMessage: this.validateBorderValue.bind(this),
                }),
                PropertyPaneDropdown('alignment', {
                  label: 'Alinhamento da saudação',
                  options: [
                    { key: styles.textLeft, text: 'Esquerda' },
                    { key: styles.textCenter, text: 'Centro' },
                    { key: styles.textRight, text: 'Direita' }
                  ],
                  selectedKey: styles.textLeft
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
