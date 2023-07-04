import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'HelloWorldWebPartStrings';
import { SPComponentLoader } from '@microsoft/sp-loader';

export interface IHelloWorldWebPartProps {
  description: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
  private async _fetchSharePointData(): Promise<{ [department: string]: string[] }> {
    const client = await this.context.msGraphClientFactory.getClient('3');
    const response = await client
      .api('/sites/vitsolutionseu.sharepoint.com,8b8f57f4-1451-4aed-bcaa-07be94d3c146,0b0a58f3-099f-4a1e-8fa7-6097fb9f555e/lists/1011b9bb-b822-4747-a10b-b3e1763ddcd4/items?$expand=fields&$select=id')
      .version('v1.0')
      .get();
  
    // Group the data by department
    const groupedData: { [department: string]: string[] } = {};
    for (let item of response.value) {
      const department = item.fields.Department;
      const title = item.fields.Title;
      if (!groupedData[department]) {
        groupedData[department] = [];
      }
      groupedData[department].push(title);
    }
  
    return groupedData;
  }
  
  public async render(): Promise<void> {
    const data = await this._fetchSharePointData();
  
    let departmentList = '';
  for (let department in data) {
    departmentList += `<ul><h1>${department}</h1>`;
    for(let title of data[department]){
      departmentList += `<li>${title}</li>`;
    }
    departmentList += `</ul>`;
  }
  
    this.domElement.innerHTML = `
    <section>
    <div class="container">
      <div class="box level-1">
        <a href="#">
          <ul>
            <h1>Directie</h1>
            <li>Jeroen de bonth</li>
          </ul>
        </a>
      </div>
      <ol class="level-2-wrapper">
        <li>
          <section class="box2 level-5">
            <a href="https://vitsolutionseu.sharepoint.com/sites/Marketing" target="_blank">
             ${departmentList}
            </a>
          </section>
        </li>
            <li>
              <section class="box2 level-2 rectangle">
                <a href="https://vitsolutionseu.sharepoint.com/sites/FinanceControle" target="_blank">
                  <ul>
                    <h1>Finance & Controle</h1>
                    <li>Patrice Riebergen</li>
                    <li>sander vreugdenhil</li>
                    <li>Jan van Son</li>
                  </ul>
                </a>
              </section>
              <ol class="level-3-wrapper">
                <li>
                  <section class="level-3 rectangle">
                    <a href="#">
                      <ul>
                        <h3>Finaciel Admin</h3>
                        <li>Doin van rooij</li>
                      </ul>
                    </a>
                  </section>
                </li>
                <li>
                  <section class="level-3 rectangle">
                    <a href="#">
                      <ul>
                        <h3>Projectcontrol</h3>
                        <li>Sjoerd de bonth</li>
                        <li>Jurre Scheij</li>
                      </ul>
                    </a>
                  </section>
                </li>
              </ol>
            </li>
          </ol>
          <ol class="level-6-wrapper">
            <li>
              <section class="box2 level-5">
                <a href="https://vitsolutionseu.sharepoint.com/sites/Receptie" target="_blank">
                  <ul>
                    <h1>Receptie</h1>
                    <li>Roby Oerlemans</li>
                    <li>Marjan de Man 1/9</li>
                  </ul>
                </a>
              </section>
            </li>
          </ol>
          <ol class="level-7-wrapper">
            <li>
              <section class="box2 level-5">
                <a href="https://vitsolutionseu.sharepoint.com/sites/KAM" target="_blank">
                  <ul>
                    <h1>KAM</h1>
                    <li>Casper Goossen</li>
                  </ul>
                </a>
              </section>
            </li>
            <li>
              <section class="box2 level-5">
                <a href="https://vitsolutionseu.sharepoint.com/sites/HR373/" target="_blank">
                  <ul>
                    <h1>HR</h1>
                    <li>Geronimo Lambertus</li>
                  </ul>
                </a>
              </section>
            </li>
          </ol>
          <ol class="level-8-wrapper">
          <li>
            <section class="box2 level-5">
              <a href="#">
                <ul>
                  <h1>Verkoop</h1>
                  <li>Jos Stokman</li>
                </ul>
              </a>
            </section>
            <ol class="level-4-wrapper">
              <li>
                <a href="#"><h4 class="level-4 rectangle">Verkoop</h4></a>
              </li>
              <li>
                <a href="#"><h4 class="level-4 rectangle">Calculatie</h4></a>
              </li>
              <li>
                <a href="#"><h4 class="level-4 rectangle">Tekenkamer</h4></a>
              </li>
            </ol>
          </li>
          <li>
            <section class="box2 level-5">
              <a href="#">
                <ul>
                  <h1>Planontwikkeling</h1>
                  <li>&nbsp;</li>
                </ul>
              </a>
            </section>
            <ol class="level-4-wrapper">
              <li>
                <a href="#"> <h4 class="level-4 rectangle">Ontwikkeling</h4></a>
              </li>
              <li>
                <a href="#"><h4 class="level-4 rectangle">Takenkamer</h4></a>
              </li>
              <li>
                <a href="#"><h4 class="level-4 rectangle">Constructie</h4></a>
              </li>
            </ol>
          </li>
          <li>
            <section class="box2 level-5">
              <a href="#">
                <ul>
                  <h1>Projecten</h1>
                  <li>Bas Verbunt</li>
                </ul>
              </a>
            </section>
            <ol class="level-4-wrapper">
              <li>
                <a href="#"><h4 class="level-4 rectangle">Projectleiding</h4></a>
              </li>
              <li>
                <a href="#"
                  ><h4 class="level-4 rectangle">Wekvoorbereiding</h4></a
                >
              </li>
              <li>
                <a href="#"><h4 class="level-4 rectangle">Uivoering</h4></a>
              </li>
              <li>
                <a href="#"><h4 class="level-4 rectangle">Realisatie</h4></a>
              </li>
              <li>
                <a href="#">
                  <h4 class="level-4 rectangle">Service & Onderhoud</h4></a
                >
              </li>
            </ol>
          </li>
        </ol>
      </div>
    </section>`;
  }

  protected onInit(): Promise<void> {
    SPComponentLoader.loadCss('/sites/Seco/SiteAssets/style.css'); return super.onInit();
  }



  // private _getEnvironmentMessage(): Promise<string> {
  //   if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
  //     return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
  //       .then(context => {
  //         let environmentMessage: string = '';
  //         switch (context.app.host.name) {
  //           case 'Office': // running in Office
  //             environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
  //             break;
  //           case 'Outlook': // running in Outlook
  //             environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
  //             break;
  //           case 'Teams': // running in Teams
  //             environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
  //             break;
  //           default:
  //             throw new Error('Unknown host');
  //         }

  //         return environmentMessage;
  //       });
  //   }

  //   return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  // }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    // this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
