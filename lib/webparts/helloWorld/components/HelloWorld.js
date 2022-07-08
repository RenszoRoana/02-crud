var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
/* eslint-disable react/jsx-no-bind */
/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from "react";
import { escape } from "@microsoft/sp-lodash-subset";
import { SPHttpClient } from "@microsoft/sp-http";
import styles from "./HelloWorld.module.scss";
var HelloWorld = /** @class */ (function (_super) {
    __extends(HelloWorld, _super);
    function HelloWorld(props, state) {
        var _this = _super.call(this, props) || this;
        _this._listItemEntityTypeName = undefined;
        _this.state = {
            status: "Ready",
            items: [],
        };
        return _this;
    }
    HelloWorld.prototype._createItem = function () {
        var _this = this;
        this.setState({
            status: "Creating item...",
            items: [],
        });
        var body = JSON.stringify({
            Title: "Item ".concat(new Date().toLocaleString()),
        });
        this.props.spHttpClient
            .post("".concat(this.props.siteUrl, "/_api/web/lists/getbytitle('").concat(this.props.listName, "')/items"), SPHttpClient.configurations.v1, {
            headers: {
                Accept: "application/json;odata=nometadata",
                "Content-type": "application/json;odata=nometadata",
                "odata-version": "",
            },
            body: body,
        })
            .then(function (response) {
            return response.json();
        })
            .then(function (item) {
            _this.setState({
                status: "Item '".concat(item.Title, "' (ID: ").concat(item.Id, ") successfully created"),
                items: [],
            });
        }, function (error) {
            _this.setState({
                status: "Error while creating the item: " + error,
                items: [],
            });
        });
    };
    HelloWorld.prototype._readItem = function () {
        var _this = this;
        this.setState({
            status: 'Loading latest items...',
            items: []
        });
        this._getLatestItemId()
            .then(function (itemId) {
            if (itemId === -1) {
                throw new Error('No items found in the list');
            }
            _this.setState({
                status: "Loading information about item ID: ".concat(itemId, "..."),
                items: []
            });
            return _this.props.spHttpClient.get("".concat(_this.props.siteUrl, "/_api/web/lists/getbytitle('").concat(_this.props.listName, "')/items(").concat(itemId, ")?$select=Title,Id"), SPHttpClient.configurations.v1, {
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'odata-version': ''
                }
            });
        })
            .then(function (response) {
            return response.json();
        })
            .then(function (item) {
            _this.setState({
                status: "Item ID: ".concat(item.Id, ", Title: ").concat(item.Title),
                items: []
            });
        }, function (error) {
            _this.setState({
                status: 'Loading latest item failed with error: ' + error,
                items: []
            });
        });
    };
    HelloWorld.prototype._getLatestItemId = function () {
        var _this = this;
        return new Promise(function (resolve, reject) {
            _this.props.spHttpClient.get("".concat(_this.props.siteUrl, "/_api/web/lists/getbytitle('").concat(_this.props.listName, "')/items?$orderby=Id desc&$top=1&$select=id"), SPHttpClient.configurations.v1, {
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'odata-version': ''
                }
            })
                .then(function (response) {
                return response.json();
            }, function (error) {
                reject(error);
            })
                .then(function (response) {
                if (response.value.length === 0) {
                    resolve(-1);
                }
                else {
                    resolve(response.value[0].Id);
                }
            })
                .catch(console.log);
        });
    };
    HelloWorld.prototype._updateItem = function () {
        var _this = this;
        this.setState({
            status: 'Loading latest items...',
            items: []
        });
        var latestItemId = undefined;
        var etag = undefined;
        var listItemEntityTypeName = undefined;
        this._getListItemEntityTypeName()
            .then(function (listItemType) {
            listItemEntityTypeName = listItemType;
            return _this._getLatestItemId();
        })
            .then(function (itemId) {
            if (itemId === -1) {
                throw new Error('No items found in the list');
            }
            latestItemId = itemId;
            _this.setState({
                status: "Loading information about item ID: ".concat(latestItemId, "..."),
                items: []
            });
            return _this.props.spHttpClient.get("".concat(_this.props.siteUrl, "/_api/web/lists/getbytitle('").concat(_this.props.listName, "')/items(").concat(latestItemId, ")?$select=Id"), SPHttpClient.configurations.v1, {
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'odata-version': ''
                }
            });
        })
            .then(function (response) {
            etag = response.headers.get('ETag');
            return response.json();
        })
            .then(function (item) {
            _this.setState({
                status: "Updating item with ID: ".concat(latestItemId, "..."),
                items: []
            });
            var body = JSON.stringify({
                '__metadata': {
                    'type': listItemEntityTypeName
                },
                'Title': "Item ".concat(new Date())
            });
            return _this.props.spHttpClient.post("".concat(_this.props.siteUrl, "/_api/web/lists/getbytitle('").concat(_this.props.listName, "')/items(").concat(item.Id, ")"), SPHttpClient.configurations.v1, {
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
            .then(function (response) {
            _this.setState({
                status: "Item with ID: ".concat(latestItemId, " successfully updated"),
                items: []
            });
        }, function (error) {
            _this.setState({
                status: "Error updating item: ".concat(error),
                items: []
            });
        });
    };
    HelloWorld.prototype._getListItemEntityTypeName = function () {
        var _this = this;
        return new Promise(function (resolve, reject) {
            if (_this._listItemEntityTypeName) {
                resolve(_this._listItemEntityTypeName);
                return;
            }
            _this.props.spHttpClient.get("".concat(_this.props.siteUrl, "/_api/web/lists/getbytitle('").concat(_this.props.listName, "')?$select=ListItemEntityTypeFullName"), SPHttpClient.configurations.v1, {
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'odata-version': ''
                }
            })
                .then(function (response) {
                return response.json();
            }, function (error) {
                reject(error);
            })
                .then(function (response) {
                _this._listItemEntityTypeName = response.ListItemEntityTypeFullName;
                resolve(_this._listItemEntityTypeName);
            })
                .catch(console.log);
        });
    };
    HelloWorld.prototype._deleteItem = function () {
        var _this = this;
        if (!window.confirm('Are you sure you want to delete the latest item?')) {
            return;
        }
        this.setState({
            status: 'Loading latest items...',
            items: []
        });
        var latestItemId = undefined;
        var etag = undefined;
        this._getLatestItemId()
            .then(function (itemId) {
            if (itemId === -1) {
                throw new Error('No items found in the list');
            }
            latestItemId = itemId;
            _this.setState({
                status: "Loading information about item ID: ".concat(latestItemId, "..."),
                items: []
            });
            return _this.props.spHttpClient.get("".concat(_this.props.siteUrl, "/_api/web/lists/getbytitle('").concat(_this.props.listName, "')/items(").concat(latestItemId, ")?$select=Id"), SPHttpClient.configurations.v1, {
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'odata-version': ''
                }
            });
        })
            .then(function (response) {
            etag = response.headers.get('ETag');
            return response.json();
        })
            .then(function (item) {
            _this.setState({
                status: "Deleting item with ID: ".concat(latestItemId, "..."),
                items: []
            });
            return _this.props.spHttpClient.post("".concat(_this.props.siteUrl, "/_api/web/lists/getbytitle('").concat(_this.props.listName, "')/items(").concat(item.Id, ")"), SPHttpClient.configurations.v1, {
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'Content-type': 'application/json;odata=verbose',
                    'odata-version': '',
                    'IF-MATCH': etag,
                    'X-HTTP-Method': 'DELETE'
                }
            });
        })
            .then(function (response) {
            _this.setState({
                status: "Item with ID: ".concat(latestItemId, " successfully deleted"),
                items: []
            });
        }, function (error) {
            _this.setState({
                status: "Error deleting item: ".concat(error),
                items: []
            });
        });
    };
    HelloWorld.prototype.render = function () {
        var _this = this;
        var items = this.state.items.map(function (item, i) {
            return (React.createElement("li", { key: i },
                item.Title,
                " ( ",
                item.Id,
                " ) "));
        });
        return (React.createElement("div", { className: styles.helloWorld },
            React.createElement("div", { className: styles.container },
                React.createElement("div", { className: styles.row },
                    React.createElement("div", { className: styles.column },
                        React.createElement("span", { className: styles.title }, "Welcome to SharePoint!"),
                        React.createElement("p", { className: styles.subTitle }, "Customize SharePoint experiences using Web Parts."),
                        React.createElement("p", { className: styles.description }, escape(this.props.listName)),
                        React.createElement("div", { className: "ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ".concat(styles.row) },
                            React.createElement("div", { className: "ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1" },
                                React.createElement("a", { href: "#", className: "".concat(styles.button), onClick: this._createItem },
                                    React.createElement("span", { className: styles.label }, "Create item")),
                                React.createElement("a", { href: "#", className: "".concat(styles.button), 
                                    // onClick={this._readItem}
                                    onClick: function () { return _this._readItem(); } },
                                    React.createElement("span", { className: styles.label }, "Read item")))),
                        React.createElement("div", { className: "ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ".concat(styles.row) },
                            React.createElement("div", { className: "ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1" },
                                React.createElement("a", { href: "#", className: "".concat(styles.button), 
                                    // onClick={this._updateItem}
                                    onClick: function () { return _this._updateItem(); } },
                                    React.createElement("span", { className: styles.label }, "Update item")),
                                React.createElement("a", { href: "#", className: "".concat(styles.button), 
                                    // onClick={this._deleteItem}
                                    onClick: function () { return _this._deleteItem(); } },
                                    React.createElement("span", { className: styles.label }, "Delete item")))),
                        React.createElement("div", { className: "ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ".concat(styles.row) },
                            React.createElement("div", { className: "ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1" },
                                this.state.status,
                                React.createElement("ul", null, items))))))));
    };
    return HelloWorld;
}(React.Component));
export default HelloWorld;
//# sourceMappingURL=HelloWorld.js.map