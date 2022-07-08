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
import { useState } from "react";
import { SPHttpClient } from "@microsoft/sp-http";
var useCrud = function (_a) {
    var spHttpClient = _a.spHttpClient, siteUrl = _a.siteUrl, listName = _a.listName;
    var _b = useState({
        status: "Ready",
        items: [],
    }), list = _b[0], setList = _b[1];
    var createItem = function () {
        setList(__assign(__assign({}, list), { status: "Creating", items: [] }));
        var body = JSON.stringify({
            Title: "Item ".concat(new Date()),
        });
        spHttpClient
            .post("".concat(siteUrl, "/_api/web/lists/getbytitle('").concat(listName, "')/items"), SPHttpClient.configurations.v1, {
            headers: {
                Accept: "application/json;odata=nometadata",
                "Content-type": "application/json;odata=verbose",
                "odata-version": "",
            },
            body: body,
        })
            .then(function (response) {
            return response.json();
        })
            .then(function (item) {
            setList(__assign(__assign({}, list), { status: "Item '".concat(item.Title, "' (ID: ").concat(item.Id, ") successfully created"), items: [] }));
        })
            .catch(function (error) {
            setList(__assign(__assign({}, list), { status: "Error creating item: ".concat(error), items: [] }));
        });
    };
    return {
        list: list,
        createItem: createItem,
    };
};
export default useCrud;
//# sourceMappingURL=useCrud.js.map