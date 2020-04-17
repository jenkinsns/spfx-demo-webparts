import { ISPList } from './GetSPlistitemsnojsWebPart';

export default class MockHttpClient {

    private static _items: ISPList[] = [{ Title: 'Mock Contact Person', ContactNumber: '9840462655', CompanyName: 'Jenkins',Country: 'India'}];

    public static get(restUrl: string, options?: any): Promise<ISPList[]> {
      return new Promise<ISPList[]>((resolve) => {
            resolve(MockHttpClient._items);
        });
    }
}
