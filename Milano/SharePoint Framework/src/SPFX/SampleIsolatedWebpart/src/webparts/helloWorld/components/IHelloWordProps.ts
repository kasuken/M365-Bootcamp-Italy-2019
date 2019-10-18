import { SPHttpClient, HttpClient, MSGraphClientFactory, AadHttpClientFactory } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';

export interface IHelloWordProps {
  pageContext: PageContext;
  spHttpClient: SPHttpClient;
  httpClient: HttpClient;
  graphClientFactory: MSGraphClientFactory;
  aadHttpClientFactory: AadHttpClientFactory;
}