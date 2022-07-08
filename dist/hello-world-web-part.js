define("babab0d9-ff65-47c5-a8c4-203c81fea461_0.0.1", ["@microsoft/sp-property-pane","@microsoft/sp-lodash-subset","@microsoft/sp-core-library","@microsoft/sp-webpart-base","react","react-dom","HelloWorldWebPartStrings","@microsoft/sp-http"], function(__WEBPACK_EXTERNAL_MODULE__26ea__, __WEBPACK_EXTERNAL_MODULE_Pk8u__, __WEBPACK_EXTERNAL_MODULE_UWqr__, __WEBPACK_EXTERNAL_MODULE_br4S__, __WEBPACK_EXTERNAL_MODULE_cDcd__, __WEBPACK_EXTERNAL_MODULE_faye__, __WEBPACK_EXTERNAL_MODULE_hI9z__, __WEBPACK_EXTERNAL_MODULE_vlQI__) { return /******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, { enumerable: true, get: getter });
/******/ 		}
/******/ 	};
/******/
/******/ 	// define __esModule on exports
/******/ 	__webpack_require__.r = function(exports) {
/******/ 		if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 			Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 		}
/******/ 		Object.defineProperty(exports, '__esModule', { value: true });
/******/ 	};
/******/
/******/ 	// create a fake namespace object
/******/ 	// mode & 1: value is a module id, require it
/******/ 	// mode & 2: merge all properties of value into the ns
/******/ 	// mode & 4: return value when already ns object
/******/ 	// mode & 8|1: behave like require
/******/ 	__webpack_require__.t = function(value, mode) {
/******/ 		if(mode & 1) value = __webpack_require__(value);
/******/ 		if(mode & 8) return value;
/******/ 		if((mode & 4) && typeof value === 'object' && value && value.__esModule) return value;
/******/ 		var ns = Object.create(null);
/******/ 		__webpack_require__.r(ns);
/******/ 		Object.defineProperty(ns, 'default', { enumerable: true, value: value });
/******/ 		if(mode & 2 && typeof value != 'string') for(var key in value) __webpack_require__.d(ns, key, function(key) { return value[key]; }.bind(null, key));
/******/ 		return ns;
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = "GFLV");
/******/ })
/************************************************************************/
/******/ ({

/***/ "26ea":
/*!**********************************************!*\
  !*** external "@microsoft/sp-property-pane" ***!
  \**********************************************/
/*! no static exports found */
/***/ (function(module, exports) {

module.exports = __WEBPACK_EXTERNAL_MODULE__26ea__;

/***/ }),

/***/ "7oDG":
/*!***************************************************************************************************************************************************!*\
  !*** ./node_modules/css-loader/dist/cjs.js!./node_modules/postcss-loader/src??postcss!./lib/webparts/helloWorld/components/HelloWorld.module.css ***!
  \***************************************************************************************************************************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__(/*! ../../../../node_modules/css-loader/dist/runtime/api.js */ "JPst")(false);
// Module
exports.push([module.i, ".helloWorld_39206472 .container_39206472{max-width:700px;margin:0 auto;box-shadow:0 2px 4px 0 rgba(0,0,0,.2),0 25px 50px 0 rgba(0,0,0,.1)}.helloWorld_39206472 .row_39206472{margin:0 -8px;box-sizing:border-box;zoom:1;color:#fff;background-color:#005a9e;padding:20px}.helloWorld_39206472 .row_39206472:after,.helloWorld_39206472 .row_39206472:before{display:table;content:\"\";line-height:0}.helloWorld_39206472 .row_39206472:after{clear:both}.helloWorld_39206472 .column_39206472{position:relative;min-height:1px;padding-left:8px;padding-right:8px;box-sizing:border-box}[dir=ltr] .helloWorld_39206472 .column_39206472{float:left}[dir=rtl] .helloWorld_39206472 .column_39206472{float:right}.helloWorld_39206472 .column_39206472 .ms-Grid_39206472{padding:0}@media (min-width:640px){.helloWorld_39206472 .column_39206472{width:83.3333333333%}}@media (min-width:1024px){.helloWorld_39206472 .column_39206472{width:66.6666666667%}}@media (min-width:1024px){[dir=ltr] .helloWorld_39206472 .column_39206472{left:16.6666666667%}[dir=rtl] .helloWorld_39206472 .column_39206472{right:16.6666666667%}}@media (min-width:640px){[dir=ltr] .helloWorld_39206472 .column_39206472{left:8.3333333333%}[dir=rtl] .helloWorld_39206472 .column_39206472{right:8.3333333333%}}.helloWorld_39206472 .title_39206472{font-size:21px;font-weight:100;color:#fff}.helloWorld_39206472 .description_39206472,.helloWorld_39206472 .subTitle_39206472{font-size:17px;font-weight:300;color:#fff}.helloWorld_39206472 .button_39206472{text-decoration:none;height:32px;min-width:80px;background-color:#0078d4;border-color:#0078d4;color:#fff;outline:transparent;position:relative;font-family:Segoe UI WestEuropean,Segoe UI,-apple-system,BlinkMacSystemFont,Roboto,Helvetica Neue,sans-serif;-webkit-font-smoothing:antialiased;font-size:14px;font-weight:400;border-width:0;text-align:center;cursor:pointer;display:inline-block;padding:0 16px}.helloWorld_39206472 .button_39206472 .label_39206472{font-weight:600;font-size:14px;height:32px;line-height:32px;margin:0 4px;vertical-align:top;display:inline-block}.helloWorld_39206472 .button_39206472.disabled_39206472,.helloWorld_39206472 .button_39206472:disabled{background-color:#f4f4f4;border-color:#f4f4f4;cursor:default;pointer-events:none}.helloWorld_39206472 .button_39206472.disabled_39206472 .label_39206472,.helloWorld_39206472 .button_39206472:disabled .label_39206472{color:#a6a6a6}", ""]);


/***/ }),

/***/ "9SGU":
/*!**********************************************************!*\
  !*** ./lib/webparts/helloWorld/components/HelloWorld.js ***!
  \**********************************************************/
/*! exports provided: default */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ "cDcd");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _microsoft_sp_lodash_subset__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @microsoft/sp-lodash-subset */ "Pk8u");
/* harmony import */ var _microsoft_sp_lodash_subset__WEBPACK_IMPORTED_MODULE_1___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_lodash_subset__WEBPACK_IMPORTED_MODULE_1__);
/* harmony import */ var _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @microsoft/sp-http */ "vlQI");
/* harmony import */ var _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_2___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_http__WEBPACK_IMPORTED_MODULE_2__);
/* harmony import */ var _HelloWorld_module_scss__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./HelloWorld.module.scss */ "nen0");
var __extends = (undefined && undefined.__extends) || (function () {
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
            .post("".concat(this.props.siteUrl, "/_api/web/lists/getbytitle('").concat(this.props.listName, "')/items"), _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_2__["SPHttpClient"].configurations.v1, {
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
            return _this.props.spHttpClient.get("".concat(_this.props.siteUrl, "/_api/web/lists/getbytitle('").concat(_this.props.listName, "')/items(").concat(itemId, ")?$select=Title,Id"), _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_2__["SPHttpClient"].configurations.v1, {
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
            _this.props.spHttpClient.get("".concat(_this.props.siteUrl, "/_api/web/lists/getbytitle('").concat(_this.props.listName, "')/items?$orderby=Id desc&$top=1&$select=id"), _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_2__["SPHttpClient"].configurations.v1, {
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
            return _this.props.spHttpClient.get("".concat(_this.props.siteUrl, "/_api/web/lists/getbytitle('").concat(_this.props.listName, "')/items(").concat(latestItemId, ")?$select=Id"), _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_2__["SPHttpClient"].configurations.v1, {
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
            return _this.props.spHttpClient.post("".concat(_this.props.siteUrl, "/_api/web/lists/getbytitle('").concat(_this.props.listName, "')/items(").concat(item.Id, ")"), _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_2__["SPHttpClient"].configurations.v1, {
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
            _this.props.spHttpClient.get("".concat(_this.props.siteUrl, "/_api/web/lists/getbytitle('").concat(_this.props.listName, "')?$select=ListItemEntityTypeFullName"), _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_2__["SPHttpClient"].configurations.v1, {
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
            return _this.props.spHttpClient.get("".concat(_this.props.siteUrl, "/_api/web/lists/getbytitle('").concat(_this.props.listName, "')/items(").concat(latestItemId, ")?$select=Id"), _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_2__["SPHttpClient"].configurations.v1, {
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
            return _this.props.spHttpClient.post("".concat(_this.props.siteUrl, "/_api/web/lists/getbytitle('").concat(_this.props.listName, "')/items(").concat(item.Id, ")"), _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_2__["SPHttpClient"].configurations.v1, {
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
            return (react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("li", { key: i },
                item.Title,
                " ( ",
                item.Id,
                " ) "));
        });
        return (react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", { className: _HelloWorld_module_scss__WEBPACK_IMPORTED_MODULE_3__["default"].helloWorld },
            react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", { className: _HelloWorld_module_scss__WEBPACK_IMPORTED_MODULE_3__["default"].container },
                react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", { className: _HelloWorld_module_scss__WEBPACK_IMPORTED_MODULE_3__["default"].row },
                    react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", { className: _HelloWorld_module_scss__WEBPACK_IMPORTED_MODULE_3__["default"].column },
                        react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("span", { className: _HelloWorld_module_scss__WEBPACK_IMPORTED_MODULE_3__["default"].title }, "Welcome to SharePoint!"),
                        react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("p", { className: _HelloWorld_module_scss__WEBPACK_IMPORTED_MODULE_3__["default"].subTitle }, "Customize SharePoint experiences using Web Parts."),
                        react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("p", { className: _HelloWorld_module_scss__WEBPACK_IMPORTED_MODULE_3__["default"].description }, Object(_microsoft_sp_lodash_subset__WEBPACK_IMPORTED_MODULE_1__["escape"])(this.props.listName)),
                        react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", { className: "ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ".concat(_HelloWorld_module_scss__WEBPACK_IMPORTED_MODULE_3__["default"].row) },
                            react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", { className: "ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1" },
                                react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("a", { href: "#", className: "".concat(_HelloWorld_module_scss__WEBPACK_IMPORTED_MODULE_3__["default"].button), onClick: this._createItem },
                                    react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("span", { className: _HelloWorld_module_scss__WEBPACK_IMPORTED_MODULE_3__["default"].label }, "Create item")),
                                react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("a", { href: "#", className: "".concat(_HelloWorld_module_scss__WEBPACK_IMPORTED_MODULE_3__["default"].button), 
                                    // onClick={this._readItem}
                                    onClick: function () { return _this._readItem(); } },
                                    react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("span", { className: _HelloWorld_module_scss__WEBPACK_IMPORTED_MODULE_3__["default"].label }, "Read item")))),
                        react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", { className: "ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ".concat(_HelloWorld_module_scss__WEBPACK_IMPORTED_MODULE_3__["default"].row) },
                            react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", { className: "ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1" },
                                react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("a", { href: "#", className: "".concat(_HelloWorld_module_scss__WEBPACK_IMPORTED_MODULE_3__["default"].button), 
                                    // onClick={this._updateItem}
                                    onClick: function () { return _this._updateItem(); } },
                                    react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("span", { className: _HelloWorld_module_scss__WEBPACK_IMPORTED_MODULE_3__["default"].label }, "Update item")),
                                react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("a", { href: "#", className: "".concat(_HelloWorld_module_scss__WEBPACK_IMPORTED_MODULE_3__["default"].button), 
                                    // onClick={this._deleteItem}
                                    onClick: function () { return _this._deleteItem(); } },
                                    react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("span", { className: _HelloWorld_module_scss__WEBPACK_IMPORTED_MODULE_3__["default"].label }, "Delete item")))),
                        react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", { className: "ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ".concat(_HelloWorld_module_scss__WEBPACK_IMPORTED_MODULE_3__["default"].row) },
                            react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", { className: "ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1" },
                                this.state.status,
                                react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("ul", null, items))))))));
    };
    return HelloWorld;
}(react__WEBPACK_IMPORTED_MODULE_0__["Component"]));
/* harmony default export */ __webpack_exports__["default"] = (HelloWorld);


/***/ }),

/***/ "GFLV":
/*!******************************************************!*\
  !*** ./lib/webparts/helloWorld/HelloWorldWebPart.js ***!
  \******************************************************/
/*! exports provided: default */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ "cDcd");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var react_dom__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! react-dom */ "faye");
/* harmony import */ var react_dom__WEBPACK_IMPORTED_MODULE_1___default = /*#__PURE__*/__webpack_require__.n(react_dom__WEBPACK_IMPORTED_MODULE_1__);
/* harmony import */ var _microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @microsoft/sp-core-library */ "UWqr");
/* harmony import */ var _microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_2___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_2__);
/* harmony import */ var _microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @microsoft/sp-property-pane */ "26ea");
/* harmony import */ var _microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_3___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_3__);
/* harmony import */ var _microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! @microsoft/sp-webpart-base */ "br4S");
/* harmony import */ var _microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_4___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_4__);
/* harmony import */ var HelloWorldWebPartStrings__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! HelloWorldWebPartStrings */ "hI9z");
/* harmony import */ var HelloWorldWebPartStrings__WEBPACK_IMPORTED_MODULE_5___default = /*#__PURE__*/__webpack_require__.n(HelloWorldWebPartStrings__WEBPACK_IMPORTED_MODULE_5__);
/* harmony import */ var _components_HelloWorld__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ./components/HelloWorld */ "9SGU");
var __extends = (undefined && undefined.__extends) || (function () {
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







var HelloWorldWebPart = /** @class */ (function (_super) {
    __extends(HelloWorldWebPart, _super);
    function HelloWorldWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this._isDarkTheme = false;
        _this._environmentMessage = '';
        return _this;
    }
    HelloWorldWebPart.prototype.render = function () {
        var element = react__WEBPACK_IMPORTED_MODULE_0__["createElement"](_components_HelloWorld__WEBPACK_IMPORTED_MODULE_6__["default"], {
            listName: 'Crud',
            spHttpClient: this.context.spHttpClient,
            siteUrl: this.context.pageContext.web.absoluteUrl,
        });
        react_dom__WEBPACK_IMPORTED_MODULE_1__["render"](element, this.domElement);
    };
    HelloWorldWebPart.prototype.onInit = function () {
        this._environmentMessage = this._getEnvironmentMessage();
        return _super.prototype.onInit.call(this);
    };
    HelloWorldWebPart.prototype._getEnvironmentMessage = function () {
        if (!!this.context.sdks.microsoftTeams) { // running in Teams
            return this.context.isServedFromLocalhost ? HelloWorldWebPartStrings__WEBPACK_IMPORTED_MODULE_5__["AppLocalEnvironmentTeams"] : HelloWorldWebPartStrings__WEBPACK_IMPORTED_MODULE_5__["AppTeamsTabEnvironment"];
        }
        return this.context.isServedFromLocalhost ? HelloWorldWebPartStrings__WEBPACK_IMPORTED_MODULE_5__["AppLocalEnvironmentSharePoint"] : HelloWorldWebPartStrings__WEBPACK_IMPORTED_MODULE_5__["AppSharePointEnvironment"];
    };
    HelloWorldWebPart.prototype.onThemeChanged = function (currentTheme) {
        if (!currentTheme) {
            return;
        }
        this._isDarkTheme = !!currentTheme.isInverted;
        var semanticColors = currentTheme.semanticColors;
        if (semanticColors) {
            this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
            this.domElement.style.setProperty('--link', semanticColors.link || null);
            this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
        }
    };
    HelloWorldWebPart.prototype.onDispose = function () {
        react_dom__WEBPACK_IMPORTED_MODULE_1__["unmountComponentAtNode"](this.domElement);
    };
    Object.defineProperty(HelloWorldWebPart.prototype, "dataVersion", {
        get: function () {
            return _microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_2__["Version"].parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    HelloWorldWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: HelloWorldWebPartStrings__WEBPACK_IMPORTED_MODULE_5__["PropertyPaneDescription"]
                    },
                    groups: [
                        {
                            groupName: HelloWorldWebPartStrings__WEBPACK_IMPORTED_MODULE_5__["BasicGroupName"],
                            groupFields: [
                                Object(_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_3__["PropertyPaneTextField"])('description', {
                                    label: HelloWorldWebPartStrings__WEBPACK_IMPORTED_MODULE_5__["DescriptionFieldLabel"]
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return HelloWorldWebPart;
}(_microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_4__["BaseClientSideWebPart"]));
/* harmony default export */ __webpack_exports__["default"] = (HelloWorldWebPart);


/***/ }),

/***/ "JPst":
/*!*****************************************************!*\
  !*** ./node_modules/css-loader/dist/runtime/api.js ***!
  \*****************************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


/*
  MIT License http://www.opensource.org/licenses/mit-license.php
  Author Tobias Koppers @sokra
*/
// css base code, injected by the css-loader
// eslint-disable-next-line func-names
module.exports = function (useSourceMap) {
  var list = []; // return the list of modules as css string

  list.toString = function toString() {
    return this.map(function (item) {
      var content = cssWithMappingToString(item, useSourceMap);

      if (item[2]) {
        return "@media ".concat(item[2], "{").concat(content, "}");
      }

      return content;
    }).join('');
  }; // import a list of modules into the list
  // eslint-disable-next-line func-names


  list.i = function (modules, mediaQuery) {
    if (typeof modules === 'string') {
      // eslint-disable-next-line no-param-reassign
      modules = [[null, modules, '']];
    }

    var alreadyImportedModules = {};

    for (var i = 0; i < this.length; i++) {
      // eslint-disable-next-line prefer-destructuring
      var id = this[i][0];

      if (id != null) {
        alreadyImportedModules[id] = true;
      }
    }

    for (var _i = 0; _i < modules.length; _i++) {
      var item = modules[_i]; // skip already imported module
      // this implementation is not 100% perfect for weird media query combinations
      // when a module is imported multiple times with different media queries.
      // I hope this will never occur (Hey this way we have smaller bundles)

      if (item[0] == null || !alreadyImportedModules[item[0]]) {
        if (mediaQuery && !item[2]) {
          item[2] = mediaQuery;
        } else if (mediaQuery) {
          item[2] = "(".concat(item[2], ") and (").concat(mediaQuery, ")");
        }

        list.push(item);
      }
    }
  };

  return list;
};

function cssWithMappingToString(item, useSourceMap) {
  var content = item[1] || ''; // eslint-disable-next-line prefer-destructuring

  var cssMapping = item[3];

  if (!cssMapping) {
    return content;
  }

  if (useSourceMap && typeof btoa === 'function') {
    var sourceMapping = toComment(cssMapping);
    var sourceURLs = cssMapping.sources.map(function (source) {
      return "/*# sourceURL=".concat(cssMapping.sourceRoot).concat(source, " */");
    });
    return [content].concat(sourceURLs).concat([sourceMapping]).join('\n');
  }

  return [content].join('\n');
} // Adapted from convert-source-map (MIT)


function toComment(sourceMap) {
  // eslint-disable-next-line no-undef
  var base64 = btoa(unescape(encodeURIComponent(JSON.stringify(sourceMap))));
  var data = "sourceMappingURL=data:application/json;charset=utf-8;base64,".concat(base64);
  return "/*# ".concat(data, " */");
}

/***/ }),

/***/ "Pk8u":
/*!**********************************************!*\
  !*** external "@microsoft/sp-lodash-subset" ***!
  \**********************************************/
/*! no static exports found */
/***/ (function(module, exports) {

module.exports = __WEBPACK_EXTERNAL_MODULE_Pk8u__;

/***/ }),

/***/ "UWqr":
/*!*********************************************!*\
  !*** external "@microsoft/sp-core-library" ***!
  \*********************************************/
/*! no static exports found */
/***/ (function(module, exports) {

module.exports = __WEBPACK_EXTERNAL_MODULE_UWqr__;

/***/ }),

/***/ "br4S":
/*!*********************************************!*\
  !*** external "@microsoft/sp-webpart-base" ***!
  \*********************************************/
/*! no static exports found */
/***/ (function(module, exports) {

module.exports = __WEBPACK_EXTERNAL_MODULE_br4S__;

/***/ }),

/***/ "cDcd":
/*!************************!*\
  !*** external "react" ***!
  \************************/
/*! no static exports found */
/***/ (function(module, exports) {

module.exports = __WEBPACK_EXTERNAL_MODULE_cDcd__;

/***/ }),

/***/ "faye":
/*!****************************!*\
  !*** external "react-dom" ***!
  \****************************/
/*! no static exports found */
/***/ (function(module, exports) {

module.exports = __WEBPACK_EXTERNAL_MODULE_faye__;

/***/ }),

/***/ "hI9z":
/*!*******************************************!*\
  !*** external "HelloWorldWebPartStrings" ***!
  \*******************************************/
/*! no static exports found */
/***/ (function(module, exports) {

module.exports = __WEBPACK_EXTERNAL_MODULE_hI9z__;

/***/ }),

/***/ "nen0":
/*!**********************************************************************!*\
  !*** ./lib/webparts/helloWorld/components/HelloWorld.module.scss.js ***!
  \**********************************************************************/
/*! exports provided: default */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* tslint:disable */
__webpack_require__(/*! ./HelloWorld.module.css */ "tojx");
var styles = {
    helloWorld: 'helloWorld_39206472',
    container: 'container_39206472',
    row: 'row_39206472',
    column: 'column_39206472',
    'ms-Grid': 'ms-Grid_39206472',
    title: 'title_39206472',
    subTitle: 'subTitle_39206472',
    description: 'description_39206472',
    button: 'button_39206472',
    label: 'label_39206472',
    disabled: 'disabled_39206472'
};
/* harmony default export */ __webpack_exports__["default"] = (styles);
/* tslint:enable */ 


/***/ }),

/***/ "ruv1":
/*!*******************************************************************************************************************!*\
  !*** ./node_modules/@microsoft/loader-load-themed-styles/node_modules/@microsoft/load-themed-styles/lib/index.js ***!
  \*******************************************************************************************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
/* WEBPACK VAR INJECTION */(function(global) {
// Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
// See LICENSE in the project root for license information.
var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.splitStyles = exports.detokenize = exports.clearStyles = exports.loadTheme = exports.flush = exports.configureRunMode = exports.configureLoadStyles = exports.loadStyles = void 0;
// Store the theming state in __themeState__ global scope for reuse in the case of duplicate
// load-themed-styles hosted on the page.
var _root = typeof window === 'undefined' ? global : window; // eslint-disable-line @typescript-eslint/no-explicit-any
// Nonce string to inject into script tag if one provided. This is used in CSP (Content Security Policy).
var _styleNonce = _root && _root.CSPSettings && _root.CSPSettings.nonce;
var _themeState = initializeThemeState();
/**
 * Matches theming tokens. For example, "[theme: themeSlotName, default: #FFF]" (including the quotes).
 */
var _themeTokenRegex = /[\'\"]\[theme:\s*(\w+)\s*(?:\,\s*default:\s*([\\"\']?[\.\,\(\)\#\-\s\w]*[\.\,\(\)\#\-\w][\"\']?))?\s*\][\'\"]/g;
var now = function () {
    return typeof performance !== 'undefined' && !!performance.now ? performance.now() : Date.now();
};
function measure(func) {
    var start = now();
    func();
    var end = now();
    _themeState.perf.duration += end - start;
}
/**
 * initialize global state object
 */
function initializeThemeState() {
    var state = _root.__themeState__ || {
        theme: undefined,
        lastStyleElement: undefined,
        registeredStyles: []
    };
    if (!state.runState) {
        state = __assign(__assign({}, state), { perf: {
                count: 0,
                duration: 0
            }, runState: {
                flushTimer: 0,
                mode: 0 /* sync */,
                buffer: []
            } });
    }
    if (!state.registeredThemableStyles) {
        state = __assign(__assign({}, state), { registeredThemableStyles: [] });
    }
    _root.__themeState__ = state;
    return state;
}
/**
 * Loads a set of style text. If it is registered too early, we will register it when the window.load
 * event is fired.
 * @param {string | ThemableArray} styles Themable style text to register.
 * @param {boolean} loadAsync When true, always load styles in async mode, irrespective of current sync mode.
 */
function loadStyles(styles, loadAsync) {
    if (loadAsync === void 0) { loadAsync = false; }
    measure(function () {
        var styleParts = Array.isArray(styles) ? styles : splitStyles(styles);
        var _a = _themeState.runState, mode = _a.mode, buffer = _a.buffer, flushTimer = _a.flushTimer;
        if (loadAsync || mode === 1 /* async */) {
            buffer.push(styleParts);
            if (!flushTimer) {
                _themeState.runState.flushTimer = asyncLoadStyles();
            }
        }
        else {
            applyThemableStyles(styleParts);
        }
    });
}
exports.loadStyles = loadStyles;
/**
 * Allows for customizable loadStyles logic. e.g. for server side rendering application
 * @param {(processedStyles: string, rawStyles?: string | ThemableArray) => void}
 * a loadStyles callback that gets called when styles are loaded or reloaded
 */
function configureLoadStyles(loadStylesFn) {
    _themeState.loadStyles = loadStylesFn;
}
exports.configureLoadStyles = configureLoadStyles;
/**
 * Configure run mode of load-themable-styles
 * @param mode load-themable-styles run mode, async or sync
 */
function configureRunMode(mode) {
    _themeState.runState.mode = mode;
}
exports.configureRunMode = configureRunMode;
/**
 * external code can call flush to synchronously force processing of currently buffered styles
 */
function flush() {
    measure(function () {
        var styleArrays = _themeState.runState.buffer.slice();
        _themeState.runState.buffer = [];
        var mergedStyleArray = [].concat.apply([], styleArrays);
        if (mergedStyleArray.length > 0) {
            applyThemableStyles(mergedStyleArray);
        }
    });
}
exports.flush = flush;
/**
 * register async loadStyles
 */
function asyncLoadStyles() {
    return setTimeout(function () {
        _themeState.runState.flushTimer = 0;
        flush();
    }, 0);
}
/**
 * Loads a set of style text. If it is registered too early, we will register it when the window.load event
 * is fired.
 * @param {string} styleText Style to register.
 * @param {IStyleRecord} styleRecord Existing style record to re-apply.
 */
function applyThemableStyles(stylesArray, styleRecord) {
    if (_themeState.loadStyles) {
        _themeState.loadStyles(resolveThemableArray(stylesArray).styleString, stylesArray);
    }
    else {
        registerStyles(stylesArray);
    }
}
/**
 * Registers a set theme tokens to find and replace. If styles were already registered, they will be
 * replaced.
 * @param {theme} theme JSON object of theme tokens to values.
 */
function loadTheme(theme) {
    _themeState.theme = theme;
    // reload styles.
    reloadStyles();
}
exports.loadTheme = loadTheme;
/**
 * Clear already registered style elements and style records in theme_State object
 * @param option - specify which group of registered styles should be cleared.
 * Default to be both themable and non-themable styles will be cleared
 */
function clearStyles(option) {
    if (option === void 0) { option = 3 /* all */; }
    if (option === 3 /* all */ || option === 2 /* onlyNonThemable */) {
        clearStylesInternal(_themeState.registeredStyles);
        _themeState.registeredStyles = [];
    }
    if (option === 3 /* all */ || option === 1 /* onlyThemable */) {
        clearStylesInternal(_themeState.registeredThemableStyles);
        _themeState.registeredThemableStyles = [];
    }
}
exports.clearStyles = clearStyles;
function clearStylesInternal(records) {
    records.forEach(function (styleRecord) {
        var styleElement = styleRecord && styleRecord.styleElement;
        if (styleElement && styleElement.parentElement) {
            styleElement.parentElement.removeChild(styleElement);
        }
    });
}
/**
 * Reloads styles.
 */
function reloadStyles() {
    if (_themeState.theme) {
        var themableStyles = [];
        for (var _i = 0, _a = _themeState.registeredThemableStyles; _i < _a.length; _i++) {
            var styleRecord = _a[_i];
            themableStyles.push(styleRecord.themableStyle);
        }
        if (themableStyles.length > 0) {
            clearStyles(1 /* onlyThemable */);
            applyThemableStyles([].concat.apply([], themableStyles));
        }
    }
}
/**
 * Find theme tokens and replaces them with provided theme values.
 * @param {string} styles Tokenized styles to fix.
 */
function detokenize(styles) {
    if (styles) {
        styles = resolveThemableArray(splitStyles(styles)).styleString;
    }
    return styles;
}
exports.detokenize = detokenize;
/**
 * Resolves ThemingInstruction objects in an array and joins the result into a string.
 * @param {ThemableArray} splitStyleArray ThemableArray to resolve and join.
 */
function resolveThemableArray(splitStyleArray) {
    var theme = _themeState.theme;
    var themable = false;
    // Resolve the array of theming instructions to an array of strings.
    // Then join the array to produce the final CSS string.
    var resolvedArray = (splitStyleArray || []).map(function (currentValue) {
        var themeSlot = currentValue.theme;
        if (themeSlot) {
            themable = true;
            // A theming annotation. Resolve it.
            var themedValue = theme ? theme[themeSlot] : undefined;
            var defaultValue = currentValue.defaultValue || 'inherit';
            // Warn to console if we hit an unthemed value even when themes are provided, but only if "DEBUG" is true.
            // Allow the themedValue to be undefined to explicitly request the default value.
            if (theme &&
                !themedValue &&
                console &&
                !(themeSlot in theme) &&
                "boolean" !== 'undefined' &&
                true) {
                console.warn("Theming value not provided for \"".concat(themeSlot, "\". Falling back to \"").concat(defaultValue, "\"."));
            }
            return themedValue || defaultValue;
        }
        else {
            // A non-themable string. Preserve it.
            return currentValue.rawString;
        }
    });
    return {
        styleString: resolvedArray.join(''),
        themable: themable
    };
}
/**
 * Split tokenized CSS into an array of strings and theme specification objects
 * @param {string} styles Tokenized styles to split.
 */
function splitStyles(styles) {
    var result = [];
    if (styles) {
        var pos = 0; // Current position in styles.
        var tokenMatch = void 0;
        while ((tokenMatch = _themeTokenRegex.exec(styles))) {
            var matchIndex = tokenMatch.index;
            if (matchIndex > pos) {
                result.push({
                    rawString: styles.substring(pos, matchIndex)
                });
            }
            result.push({
                theme: tokenMatch[1],
                defaultValue: tokenMatch[2] // May be undefined
            });
            // index of the first character after the current match
            pos = _themeTokenRegex.lastIndex;
        }
        // Push the rest of the string after the last match.
        result.push({
            rawString: styles.substring(pos)
        });
    }
    return result;
}
exports.splitStyles = splitStyles;
/**
 * Registers a set of style text. If it is registered too early, we will register it when the
 * window.load event is fired.
 * @param {ThemableArray} styleArray Array of IThemingInstruction objects to register.
 * @param {IStyleRecord} styleRecord May specify a style Element to update.
 */
function registerStyles(styleArray) {
    if (typeof document === 'undefined') {
        return;
    }
    var head = document.getElementsByTagName('head')[0];
    var styleElement = document.createElement('style');
    var _a = resolveThemableArray(styleArray), styleString = _a.styleString, themable = _a.themable;
    styleElement.setAttribute('data-load-themed-styles', 'true');
    if (_styleNonce) {
        styleElement.setAttribute('nonce', _styleNonce);
    }
    styleElement.appendChild(document.createTextNode(styleString));
    _themeState.perf.count++;
    head.appendChild(styleElement);
    var ev = document.createEvent('HTMLEvents');
    ev.initEvent('styleinsert', true /* bubbleEvent */, false /* cancelable */);
    ev.args = {
        newStyle: styleElement
    };
    document.dispatchEvent(ev);
    var record = {
        styleElement: styleElement,
        themableStyle: styleArray
    };
    if (themable) {
        _themeState.registeredThemableStyles.push(record);
    }
    else {
        _themeState.registeredStyles.push(record);
    }
}
//# sourceMappingURL=index.js.map
/* WEBPACK VAR INJECTION */}.call(this, __webpack_require__(/*! ./../../../../../../webpack/buildin/global.js */ "yLpj")))

/***/ }),

/***/ "tojx":
/*!******************************************************************!*\
  !*** ./lib/webparts/helloWorld/components/HelloWorld.module.css ***!
  \******************************************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

var content = __webpack_require__(/*! !../../../../node_modules/css-loader/dist/cjs.js!../../../../node_modules/postcss-loader/src??postcss!./HelloWorld.module.css */ "7oDG");
var loader = __webpack_require__(/*! ./node_modules/@microsoft/loader-load-themed-styles/node_modules/@microsoft/load-themed-styles/lib/index.js */ "ruv1");

if(typeof content === "string") content = [[module.i, content]];

// add the styles to the DOM
for (var i = 0; i < content.length; i++) loader.loadStyles(content[i][1], true);

if(content.locals) module.exports = content.locals;

/***/ }),

/***/ "vlQI":
/*!*************************************!*\
  !*** external "@microsoft/sp-http" ***!
  \*************************************/
/*! no static exports found */
/***/ (function(module, exports) {

module.exports = __WEBPACK_EXTERNAL_MODULE_vlQI__;

/***/ }),

/***/ "yLpj":
/*!***********************************!*\
  !*** (webpack)/buildin/global.js ***!
  \***********************************/
/*! no static exports found */
/***/ (function(module, exports) {

var g;

// This works in non-strict mode
g = (function() {
	return this;
})();

try {
	// This works if eval is allowed (see CSP)
	g = g || new Function("return this")();
} catch (e) {
	// This works if the window reference is available
	if (typeof window === "object") g = window;
}

// g can still be undefined, but nothing to do about it...
// We return undefined, instead of nothing here, so it's
// easier to handle this case. if(!global) { ...}

module.exports = g;


/***/ })

/******/ })});;
//# sourceMappingURL=hello-world-web-part.js.map