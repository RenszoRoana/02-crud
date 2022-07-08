/* eslint-disable react/jsx-no-bind */
/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from "react";

import { escape } from "@microsoft/sp-lodash-subset";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

import {
  IHelloWorldProps,
  IHelloWorldState,
  IListItem,
} from "../interfaces/IHelloWorldProps";
import styles from "./HelloWorld.module.scss";

export default class HelloWorld extends React.Component<
  IHelloWorldProps,
  IHelloWorldState
> {
  private _listItemEntityTypeName: string = undefined;
  public constructor(props: IHelloWorldProps, state: IHelloWorldState) {
    super(props);

    this.state = {
      status: "Ready",
      items: [],
    };
  }

  private _createItem(): void {
    this.setState({
      status: "Creating item...",
      items: [],
    });

    const body: string = JSON.stringify({
      Title: `Item ${new Date().toLocaleString()}`,
    });

    this.props.spHttpClient
      .post(
        `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
            "Content-type": "application/json;odata=nometadata",
            "odata-version": "",
          },
          body: body,
        }
      )
      .then((response: SPHttpClientResponse): Promise<IListItem> => {
        return response.json();
      })
      .then(
        (item: IListItem): void => {
          this.setState({
            status: `Item '${item.Title}' (ID: ${item.Id}) successfully created`,
            items: [],
          });
        },
        (error: any): void => {
          this.setState({
            status: "Error while creating the item: " + error,
            items: [],
          });
        }
      );
  }

  private _readItem(): void {
    this.setState({
      status: 'Loading latest items...',
      items: []
    });
    this._getLatestItemId()
      .then((itemId: number): Promise<SPHttpClientResponse> => {
        if (itemId === -1) {
          throw new Error('No items found in the list');
        }

        this.setState({
          status: `Loading information about item ID: ${itemId}...`,
          items: []
        });
        return this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${itemId})?$select=Title,Id`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          });
      })
      .then((response: SPHttpClientResponse): Promise<IListItem> => {
        return response.json();
      })
      .then((item: IListItem): void => {
        this.setState({
          status: `Item ID: ${item.Id}, Title: ${item.Title}`,
          items: []
        });
      }, (error: any): void => {
        this.setState({
          status: 'Loading latest item failed with error: ' + error,
          items: []
        });
      });
  }

  private _getLatestItemId(): Promise<number> {
      return new Promise<number>((resolve: (itemId: number) => void, reject: (error: any) => void): void => {
        this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items?$orderby=Id desc&$top=1&$select=id`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          })
          .then((response: SPHttpClientResponse): Promise<{ value: { Id: number }[] }> => {
            return response.json();
          }, (error: any): void => {
            reject(error);
          })
          .then((response: { value: { Id: number }[] }): void => {
            if (response.value.length === 0) {
              resolve(-1);
            }
            else {
              resolve(response.value[0].Id);
            }
          })
          .catch(console.log)
      });
    } 

  private _updateItem(): void {
    this.setState({
      status: 'Loading latest items...',
      items: []
    });
    let latestItemId: number = undefined;
    let etag: string = undefined;
    let listItemEntityTypeName: string = undefined;
    this._getListItemEntityTypeName()
      .then((listItemType: string): Promise<number> => {
        listItemEntityTypeName = listItemType;
        return this._getLatestItemId();
      })
      .then((itemId: number): Promise<SPHttpClientResponse> => {
        if (itemId === -1) {
          throw new Error('No items found in the list');
        }

        latestItemId = itemId;
        this.setState({
          status: `Loading information about item ID: ${latestItemId}...`,
          items: []
        });
        return this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${latestItemId})?$select=Id`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          });
      })
      .then((response: SPHttpClientResponse): Promise<IListItem> => {
        etag = response.headers.get('ETag');
        return response.json();
      })
      .then((item: IListItem): Promise<SPHttpClientResponse> => {
        this.setState({
          status: `Updating item with ID: ${latestItemId}...`,
          items: []
        });
        const body: string = JSON.stringify({
          '__metadata': {
            'type': listItemEntityTypeName
          },
          'Title': `Item ${new Date()}`
        });
        return this.props.spHttpClient.post(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${item.Id})`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-type': 'application/json;odata=verbose',
              'odata-version': '',
              'IF-MATCH': etag,
              'X-HTTP-Method': 'MERGE'
            },
            body: body
          });
      })
      .then((response: SPHttpClientResponse): void => {
        this.setState({
          status: `Item with ID: ${latestItemId} successfully updated`,
          items: []
        });
      }, (error: any): void => {
        this.setState({
          status: `Error updating item: ${error}`,
          items: []
        });
      });
  }

  private _getListItemEntityTypeName(): Promise<string> {
    return new Promise<string>((resolve: (listItemEntityTypeName: string) => void, reject: (error: any) => void): void => {
      if (this._listItemEntityTypeName) {
        resolve(this._listItemEntityTypeName);
        return;
      }

      this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')?$select=ListItemEntityTypeFullName`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        })
        .then((response: SPHttpClientResponse): Promise<{ ListItemEntityTypeFullName: string }> => {
          return response.json();
        }, (error: any): void => {
          reject(error);
        })
        .then((response: { ListItemEntityTypeFullName: string }): void => {
          this._listItemEntityTypeName = response.ListItemEntityTypeFullName;
          resolve(this._listItemEntityTypeName);
        })
        .catch(console.log);
    });
  }



  private _deleteItem(): void {
    if (!window.confirm('Are you sure you want to delete the latest item?')) {
      return;
    }

    this.setState({
      status: 'Loading latest items...',
      items: []
    });
    let latestItemId: number = undefined;
    let etag: string = undefined;
    this._getLatestItemId()
      .then((itemId: number): Promise<SPHttpClientResponse> => {
        if (itemId === -1) {
          throw new Error('No items found in the list');
        }

        latestItemId = itemId;
        this.setState({
          status: `Loading information about item ID: ${latestItemId}...`,
          items: []
        });
        return this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${latestItemId})?$select=Id`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          });
      })
      .then((response: SPHttpClientResponse): Promise<IListItem> => {
        etag = response.headers.get('ETag');
        return response.json();
      })
      .then((item: IListItem): Promise<SPHttpClientResponse> => {
        this.setState({
          status: `Deleting item with ID: ${latestItemId}...`,
          items: []
        });
        return this.props.spHttpClient.post(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${item.Id})`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-type': 'application/json;odata=verbose',
              'odata-version': '',
              'IF-MATCH': etag,
              'X-HTTP-Method': 'DELETE'
            }
          });
      })
      .then((response: SPHttpClientResponse): void => {
        this.setState({
          status: `Item with ID: ${latestItemId} successfully deleted`,
          items: []
        });
      }, (error: any): void => {
        this.setState({
          status: `Error deleting item: ${error}`,
          items: []
        });
      });
  }

  public render(): React.ReactElement<IHelloWorldProps> {
    const items: JSX.Element[] = this.state.items.map((item: IListItem, i: number): JSX.Element => {  
      return (  
        <li key={i}>{item.Title} ( {item.Id} ) </li>  
      );  
    });  

    return (
      <div className={styles.helloWorld}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to SharePoint!</span>
              <p className={styles.subTitle}>
                Customize SharePoint experiences using Web Parts.
              </p>
              <p className={styles.description}>
                {escape(this.props.listName)}
              </p>

              <div
                className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}`}
              >
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <a
                    href="#"
                    className={`${styles.button}`}
                    onClick={this._createItem}
                    // onClick={() => this._createItem()}
                  >
                    <span className={styles.label}>Create item</span>
                  </a>
                  <a
                    href="#"
                    className={`${styles.button}`}
                    // onClick={this._readItem}
                    onClick={() => this._readItem()}
                  >
                    <span className={styles.label}>Read item</span>
                  </a>
                </div>
              </div>

              <div
                className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}`}
              >
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <a
                    href="#"
                    className={`${styles.button}`}
                    // onClick={this._updateItem}
                    onClick={() => this._updateItem()}
                  >
                    <span className={styles.label}>Update item</span>
                  </a>
                  <a
                    href="#"
                    className={`${styles.button}`}
                    // onClick={this._deleteItem}
                    onClick={() => this._deleteItem()}
                  >
                    <span className={styles.label}>Delete item</span>
                  </a>
                </div>
              </div>

              <div
                className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}`}
              >
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  {this.state.status}
                  <ul>{items}</ul>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
