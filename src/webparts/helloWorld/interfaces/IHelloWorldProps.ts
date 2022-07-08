export interface IHelloWorldProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}

export interface IListItem {
  Title?: string;
  id: number;
}

export interface IHelloWorldState {
  status: string;
  items: IListItem[];
}