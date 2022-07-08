import { SPHttpClient } from '@microsoft/sp-http';
export interface ICRUDProps {
    listName: string;
    spHttpClient: SPHttpClient;
    siteUrl: string;
}
export interface ICRUDState {
    status: string;
    items: IListItem[];
}
export interface IListItem {
    Title?: string;
    Id: number;
}
//# sourceMappingURL=crud.interface.d.ts.map