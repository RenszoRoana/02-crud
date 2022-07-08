import { SPHttpClient } from '@microsoft/sp-http'; 

export interface IHelloWorldProps {
  listName: string;
  spHttpClient: SPHttpClient;
  siteUrl: string;
}

export interface IListItem {
  Title?: string;
  Id: number;
}

export interface IHelloWorldState {
  status: string;
  items: IListItem[];
}