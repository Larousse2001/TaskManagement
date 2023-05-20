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
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable react/self-closing-comp */
import * as React from "react";
import styles from "./Register.module.scss";
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
var Register = /** @class */ (function (_super) {
    __extends(Register, _super);
    function Register() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        //Create Item
        _this.createItem = function () { return __awaiter(_this, void 0, void 0, function () {
            var addItem, e_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, sp.web.lists.getByTitle("Employees").items.add({
                                Fullname: document.getElementById("fullName")
                                    .value,
                                Email: document.getElementById("email").value,
                                Team: document.getElementById("team").value,
                                Role: "Team Lead",
                            })];
                    case 1:
                        addItem = _a.sent();
                        console.log(addItem);
                        alert("Item created successfully with ID: ".concat(addItem.data.ID));
                        return [3 /*break*/, 3];
                    case 2:
                        e_1 = _a.sent();
                        console.error(e_1);
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        }); };
        return _this;
    }
    Register.prototype.render = function () {
        return (React.createElement("div", { className: styles.register },
            React.createElement("h1", { className: styles.title }, "Registration for Team Lead \uD83D\uDC54"),
            React.createElement("div", { className: styles.teamOv },
                React.createElement("form", { onSubmit: function (event) { return event.preventDefault(); } },
                    React.createElement("div", { className: styles.itemField },
                        React.createElement("div", { className: styles.fieldLabel }, "Full Name"),
                        React.createElement("input", { type: "text", id: "fullName" })),
                    React.createElement("div", { className: styles.itemField },
                        React.createElement("div", { className: styles.fieldLabel }, "Email"),
                        React.createElement("input", { type: "text", id: "email" })),
                    React.createElement("div", { className: styles.itemField },
                        React.createElement("div", { className: styles.fieldLabel }, "Team Name:"),
                        React.createElement("input", { type: "text", id: "team" })),
                    React.createElement("div", { className: styles.buttonSection },
                        React.createElement("div", { className: styles.button },
                            React.createElement("span", { className: styles.label, onClick: this.createItem }, "Register")),
                        React.createElement("div", { className: styles.button },
                            React.createElement("span", { className: styles.label }, "Cancel")))))));
    };
    return Register;
}(React.Component));
export default Register;
//# sourceMappingURL=Register.js.map