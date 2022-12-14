import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { ISPHttpClientOptions , SPHttpClient, SPHttpClientResponse  } from '@microsoft/sp-http';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
// import { IReadonlyTheme } from '@microsoft/sp-component-base';
// import { escape } from '@microsoft/sp-lodash-subset';

// import styles from './IpmDocUploderWebPart.module.scss';
import * as strings from 'IpmDocUploderWebPartStrings';


export interface IIpmDocUploderWebPartProps {
  description: string;
}

export default class IpmDocUploderWebPart extends BaseClientSideWebPart<IIpmDocUploderWebPartProps> {

  // private _isDarkTheme: boolean = false;
  // private _environmentMessage: string = '';

  public render(): void {
    this.domElement.innerHTML = `
     <input type="file" id="uploadFile" multiple="true" value="Upload File" />
     <input type="button" id="uploadButton" value="Upload" />
    `;
    this.bindUploadButton();
  }

private bindUploadButton(): void {
  this.domElement.querySelector('#uploadButton').addEventListener('click', ()=> {this.UploadFiles(); });
}
private UploadFiles(): void{
  const files = (<HTMLInputElement>document.getElementById('uploadFile')).files;
  for (let i = 0; i < files.length; i++){
     const file = files[i];

     const spOpts: ISPHttpClientOptions = {
      headers:{
        "Accept":"application/json",
        "Content-Type": "application/json"
      },
      body: file
     };
     const url = `${ this.context.pageContext.web.absoluteUrl}/_api/web/Lists/getByTitle('Documents')/RootFolder/Files/Add(url='${file.name}', overwrite=true)`;
   
     this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spOpts).then((response: SPHttpClientResponse) => {
      
    response.json().then((responseJSON) => {
        console.log(responseJSON.Name)
      });
     });
  }
}
  // protected onInit(): Promise<void> {
  //   return this._getEnvironmentMessage().then(message => {
  //     this._environmentMessage = message;
  //   });
  // }



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

  // protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
  //   if (!currentTheme) {
  //     return;
  //   }

  //   this._isDarkTheme = !!currentTheme.isInverted;
  //   const {
  //     semanticColors
  //   } = currentTheme;

  //   if (semanticColors) {
  //     this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
  //     this.domElement.style.setProperty('--link', semanticColors.link || null);
  //     this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
  //   }

  // }

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
