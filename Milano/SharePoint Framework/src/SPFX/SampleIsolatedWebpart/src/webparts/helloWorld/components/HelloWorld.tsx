import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { SPHttpClient, SPHttpClientResponse, HttpClient, HttpClientResponse, MSGraphClient, AadHttpClient } from '@microsoft/sp-http';
import { IHelloWordProps } from './IHelloWordProps';

export interface IHelloWorldState {
  PostBody: string;
  Lists: IList[];
  Me: IMe;
  ServerDate: string;
  ErrorMessage: string;
}

export interface IList {
  Title: string;
  Id: string;
}

export interface IMe {
  Firstname: string;
  Lastname: string;
  Email: string;
}

export default class HelloWorld extends React.Component<IHelloWordProps, IHelloWorldState> {

  constructor(props) {
    super(props);    
    let meNull : IMe = { Firstname: "", Lastname: "", Email: "" };
    this.state = { PostBody: "", Lists: [], Me: meNull, ServerDate: "", ErrorMessage: ""};
  }

  // reference: https://docs.microsoft.com/en-us/sharepoint/dev/spfx/connect-to-anonymous-apis
  private callThirdPartyAPIs(): void {
    let options = { headers: [['accept', 'application/json']] };
    let apiURL = "https://jsonplaceholder.typicode.com/posts/1";
    this.props.httpClient.get(apiURL, HttpClient.configurations.v1, options).then((res: HttpClientResponse): Promise<any> => {
        return res.json();
      }).then((response: any): void => {
        this.setState({ PostBody: response.body });
      });
  }

  //reference: https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/connect-to-sharepoint
  private callSPRESTAPIs(): void {
    let siteUrl = this.props.pageContext.web.absoluteUrl;
    this.props.spHttpClient.get(`${siteUrl}/_api/web/lists?$select=Title,Id&$top=5`, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse): Promise<IList[]> => {
          if (response.ok) {
            return response.json();
          }
          else {
            this.setState({ ErrorMessage: response.statusText });
          }
        }
      ).then((result: any): void => {        
          this.setState({ Lists: result.value });
      }).catch((error: any): void => {
          this.setState({ ErrorMessage: error });
      });
  }

  //reference: https://docs.microsoft.com/en-us/sharepoint/dev/spfx/use-msgraph
  private callGraphAPIs(): void {
    this.props.graphClientFactory.getClient()
    .then((client: MSGraphClient): void => {
      client
        .api("me")
        .version("beta")
        .select("givenName,surname,mail")
        .get((error, response) => {
          if(error) {
            this.setState({ ErrorMessage: error.message });
          }
          this.setState({ 
            Me: {
              Email: response.mail,
              Firstname: response.givenName,
              Lastname: response.surname            
            }
          });
        });
      }); 
  }

  //reference: https://docs.microsoft.com/en-us/sharepoint/dev/spfx/use-aad-tutorial
  private callEnterpriseAPIs(): void {
    let aadAppClientId = "097bec0f-036b-4a90-8db8-e459fe11ffe0";
    let apiUrl = "https://d4s-bootcamp2019-api.azurewebsites.net/api/GetTime";
    this.props.aadHttpClientFactory
    .getClient(aadAppClientId)
    .then((client: AadHttpClient) => {
      return client.get(apiUrl,AadHttpClient.configurations.v1);
    })
    .then(response => {
      return response.text();
    })
    .then(json => {  
      this.setState({ ServerDate: json });    
    })
    .catch(error => {
      this.setState({ ErrorMessage: error.message });
    });
  }

  public render(): React.ReactElement<{}> {
    let call1Render = this.state.PostBody !== '' ? <p>{this.state.PostBody}</p> : null;
    let call2Render = this.state.Lists.length > 0 ? <ul>
    {
      this.state.Lists.map(list => {
        return <li>{list.Title}</li>;
      })
    }
    </ul> : null;
    let call3Render = this.state.Me.Firstname !== '' ? <p>{this.state.Me.Firstname} {this.state.Me.Lastname}<br />{this.state.Me.Email}</p> : null;
    let call4Render = this.state.ServerDate !== '' ? <p>{this.state.ServerDate}</p> : null;

    return (
      <div className={ styles.helloWorld }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <p>
                <span className={ styles.title }>Call APIs from SPFx ISOLATED webpart</span>                                       
              </p>
              <button className={ styles.button } onClick={() => this.callThirdPartyAPIs()}>
                <span className={ styles.label }>Third party APIs</span>
              </button>
              <button className={ styles.button } onClick={() => this.callSPRESTAPIs()}>
                <span className={ styles.label }>SharePoint REST APIs</span>
              </button>
              <button className={ styles.button } onClick={() => this.callGraphAPIs()}>
                <span className={ styles.label }>Graph APIs</span>
              </button>
              <button className={ styles.button } onClick={() => this.callEnterpriseAPIs()}>
                <span className={ styles.label }>Enterprise APIs</span>
              </button>
            </div>
          </div>
          <div className={ styles.row }>
            <div className={ styles.column }>
              {call1Render}
              {call2Render}
              {call3Render}
              {call4Render}
            </div>
          </div>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <p>
                {this.state.ErrorMessage}
              </p>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
