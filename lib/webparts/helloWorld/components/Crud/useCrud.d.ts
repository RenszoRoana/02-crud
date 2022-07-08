import { ICRUDProps, ICRUDState } from "./crud.interface";
interface IUseCrudProps {
    list: ICRUDState;
    createItem: () => void;
}
declare const useCrud: ({ spHttpClient, siteUrl, listName }: ICRUDProps) => IUseCrudProps;
export default useCrud;
//# sourceMappingURL=useCrud.d.ts.map