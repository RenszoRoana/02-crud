/* eslint-disable react/jsx-no-bind */
import * as React from "react";
import { escape } from "@microsoft/sp-lodash-subset";
import useCrud from "./useCrud";
import styles from "./crud.module.scss";
var CrudSPFx = function (_a) {
    var listName = _a.listName, siteUrl = _a.siteUrl, spHttpClient = _a.spHttpClient;
    var _b = useCrud({ spHttpClient: spHttpClient, siteUrl: siteUrl, listName: listName }), list = _b.list, createItem = _b.createItem;
    return (React.createElement("div", { className: styles.helloWorld },
        React.createElement("div", { className: styles.container },
            React.createElement("div", { className: styles.row },
                React.createElement("div", { className: styles.column },
                    React.createElement("span", { className: styles.title }, "Welcome to SharePoint!"),
                    React.createElement("p", { className: styles.subTitle }, "Customize SharePoint experiences using Web Parts."),
                    React.createElement("p", { className: styles.description }, escape(listName)),
                    React.createElement("div", { className: "ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ".concat(styles.row) },
                        React.createElement("div", { className: "ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1" },
                            React.createElement("a", { href: "#", className: "".concat(styles.button), onClick: function () { return createItem(); } },
                                React.createElement("span", { className: styles.label }, "Create item")))),
                    React.createElement("div", { className: "ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ".concat(styles.row) },
                        React.createElement("div", { className: "ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1" })),
                    React.createElement("div", { className: "ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ".concat(styles.row) },
                        React.createElement("div", { className: "ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1" }, list)))))));
};
export default CrudSPFx;
//# sourceMappingURL=index.js.map