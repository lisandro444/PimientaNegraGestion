import { spfi, SPFx, SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { PageContext } from '@microsoft/sp-page-context';


export class SharePointService {
  private readonly sp: SPFI;

  constructor(pageContext: PageContext) {
    this.sp = spfi().using(SPFx({ pageContext }));
  }

  public async getListItems<T>(listTitle: string, selectFields: string[], top: number = 200): Promise<T[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(listTitle)
        .items.select(...selectFields)
        .top(top)();

      return items as T[];
    } catch (err) {
      console.error(`Error fetching items from ${listTitle}:`, err);
      throw err;
    }
  }

  public async getListItemsFiltered<T>(listTitle: string, selectFields: string[], filter: string, top: number = 200): Promise<T[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(listTitle)
        .items.select(...selectFields)
        .filter(filter)
        .top(top)();

      return items as T[];
    } catch (err) {
      console.error(`Error fetching filtered items from ${listTitle}:`, err);
      throw err;
    }
  }

  public async addListItem<T>(listTitle: string, payload: Record<string, unknown>, selectFields?: string[]): Promise<T> {
    try {
      const result = await this.sp.web.lists.getByTitle(listTitle).items.add(payload);

      if (selectFields && selectFields.length > 0) {
        const createdItem = await result.item.select(...selectFields)();
        return createdItem as T;
      }

      return result.data as T;
    } catch (err) {
      console.error(`Error creating item in ${listTitle}:`, err);
      throw err;
    }
  }

  public async updateListItem(listTitle: string, id: number, payload: Record<string, unknown>): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(listTitle).items.getById(id).update(payload);
    } catch (err) {
      console.error(`Error updating item ${id} in ${listTitle}:`, err);
      throw err;
    }
  }

  public async deleteListItem(listTitle: string, id: number): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(listTitle).items.getById(id).delete();
    } catch (err) {
      console.error(`Error deleting item ${id} from ${listTitle}:`, err);
      throw err;
    }
  }
}
