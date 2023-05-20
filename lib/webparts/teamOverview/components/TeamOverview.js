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
/* eslint-disable react/no-unescaped-entities */
/* eslint-disable no-void */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable react/self-closing-comp */
import * as React from "react";
import styles from "./TeamOverview.module.scss";
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
var TeamOverview = /** @class */ (function (_super) {
    __extends(TeamOverview, _super);
    function TeamOverview(props) {
        var _this = _super.call(this, props) || this;
        _this.handleButtonClick = function () {
            _this.setState({
                selectedButton: true, // Open popup when button is clicked
            });
        };
        _this.handleButtonClose = function () {
            _this.setState({
                selectedButton: null, // Close popup when close button is clicked
            });
        };
        // Get all items
        _this.getAllItems = function () { return __awaiter(_this, void 0, void 0, function () {
            var currentTeam, items, html_1, e_1;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        return [4 /*yield*/, this.getTeamForCurrentUser()];
                    case 1:
                        currentTeam = _a.sent();
                        return [4 /*yield*/, sp.web.lists
                                .getByTitle("Employees")
                                .items.filter("Team eq '".concat(currentTeam, "'"))
                                .get()];
                    case 2:
                        items = _a.sent();
                        if (items.length > 0) {
                            html_1 = "\n        <table style=\"width:100%;border-collapse:collapse;\">\n          <tr style=\"border-bottom:1px solid #000;\">\n            <th style=\"text-align:left;padding:16px;\">Full Name</th>\n            <th style=\"text-align:left;padding:16px;\">Email</th>\n          </tr>\n      ";
                            items.map(function (item, index) {
                                html_1 += "\n          <tr style=\"border-bottom:1px solid #000;cursor:pointer;\">\n            <td style=\"text-align:left;padding:16px;\">".concat(item.Fullname, "</td>\n            <td style=\"text-align:left;padding:16px;\">").concat(item.Email, "</td>\n          </tr>\n        ");
                                setTimeout(function () {
                                    var tr = document.querySelector("#allItems tr:nth-child(".concat(index + 2, ")"));
                                    if (tr) {
                                        tr.addEventListener("click", _this.displayPopup.bind(_this, item));
                                    }
                                }, 100);
                            });
                            html_1 += "</table>";
                            document.getElementById("allItems").innerHTML = html_1;
                        }
                        else {
                            alert("List is empty.");
                        }
                        return [3 /*break*/, 4];
                    case 3:
                        e_1 = _a.sent();
                        console.error(e_1);
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        }); };
        _this.displayPopup = function (item) {
            var popupHtml = "\n    <div style=\"position: absolute;top: 95px;left: 50%;transform: translateX(-50%);background-color: #ffffff;border-radius: 5px;\n      box-shadow: 0 0 10px rgba(0, 0, 0, 0.2);padding: 20px;max-width: 400px;width: 100%;max-height: 80vh;overflow-y: auto;z-index: 9999;\">\n      <h2 style=\"font-size: 24px;margin-top: 0;\">Employee Details</h2>\n      <p style=\"margin-bottom: 10px;\">\n        <b>ID: </b>\n        <input\n          type=\"text\"\n          id=\"itemID\"\n          value=".concat(item.ID, "\n          disabled\n        ></input>\n      </p>\n      <p style=\"margin-bottom: 10px;\">\n        <b>Fullname: </b>\n        ").concat(item.Fullname, "\n      </p>\n      <p style=\"margin-bottom: 10px;\">\n        <b>Email: </b>\n        ").concat(item.Email, "\n      </p>\n      <p style=\"margin-bottom: 10px;\">\n        <b>Role: </b>\n        ").concat(item.Role, "\n      </p>\n      <button  id=\"deleteButton\" style=\"background-color: #0078d4;color: #ffffff;border: none;padding: 10px 15px;cursor: pointer;font-size: 16px;\n        border-radius: 5px;margin-right: 10px;\">Delete</button>\n      <button id=\"hide\" style=\"background-color: #0078d4;color: #ffffff;border: none;padding: 10px 15px;cursor: pointer;font-size: 16px;\n        border-radius: 5px;margin-right: 10px;\">Close</button>\n    </div>\n    ");
            setTimeout(function () {
                // Add event listener to the delete button
                var deleteButton = document.getElementById("deleteButton");
                if (deleteButton) {
                    deleteButton.addEventListener("click", function () {
                        void _this.deleteItem(item.ID);
                    });
                }
                // Add event listener to the delete button
                var hideButton = document.getElementById("hide");
                if (hideButton) {
                    hideButton.addEventListener("click", function () {
                        void _this.hidePopup();
                    });
                }
            }, 100);
            document.getElementById("task-details-popup").innerHTML = popupHtml;
        };
        //Delete Item
        _this.deleteItem = function (itemId) { return __awaiter(_this, void 0, void 0, function () {
            var deleteItem, e_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, sp.web.lists
                                .getByTitle("Employees")
                                .items.getById(itemId)
                                .delete()];
                    case 1:
                        deleteItem = _a.sent();
                        console.log(deleteItem);
                        alert("Item ID: ".concat(itemId, " deleted successfully!"));
                        this.hidePopup();
                        void this.getAllItems();
                        return [3 /*break*/, 3];
                    case 2:
                        e_2 = _a.sent();
                        console.error(e_2);
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        }); };
        //Hide the popup
        _this.hidePopup = function () {
            var popupHtml = "<div></div>";
            document.getElementById("task-details-popup").innerHTML = popupHtml;
        };
        // Get all items and return a list filled with employees' emails or an empty list
        _this.getListEmployees = function () { return __awaiter(_this, void 0, void 0, function () {
            var items, employeesEmails_1, e_3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, sp.web.lists
                                .getByTitle("Employees")
                                .items.get()];
                    case 1:
                        items = _a.sent();
                        employeesEmails_1 = [];
                        if (items.length > 0) {
                            // Loop through the retrieved items and extract the "Fullname" field value
                            items.forEach(function (item) {
                                // Add the "Fullname" field value to the employees' full names array
                                employeesEmails_1.push(item.Email);
                            });
                        }
                        console.log(employeesEmails_1);
                        // Return the employees' full names array
                        return [2 /*return*/, employeesEmails_1];
                    case 2:
                        e_3 = _a.sent();
                        console.error(e_3);
                        // Return an empty array in case of any error
                        return [2 /*return*/, []];
                    case 3: return [2 /*return*/];
                }
            });
        }); };
        // Function to get the current Team variable of the current user from SharePoint
        _this.getTeamForCurrentUser = function () { return __awaiter(_this, void 0, void 0, function () {
            var currentUser, userEmail, list, items, currentTeam, e_4;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        return [4 /*yield*/, sp.web.currentUser()];
                    case 1:
                        currentUser = _a.sent();
                        userEmail = currentUser.Email;
                        list = sp.web.lists.getByTitle("Employees");
                        return [4 /*yield*/, list.items.filter("Email eq '".concat(userEmail, "'")).get()];
                    case 2:
                        items = _a.sent();
                        if (items.length > 0) {
                            currentTeam = items[0].Team;
                            return [2 /*return*/, currentTeam];
                        }
                        else {
                            return [2 /*return*/, ""];
                        }
                        return [3 /*break*/, 4];
                    case 3:
                        e_4 = _a.sent();
                        console.error(e_4);
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        }); };
        //Create Item
        _this.createItem = function () { return __awaiter(_this, void 0, void 0, function () {
            var currentTeam, email_1, employeesEmails, existingItems, addItem, e_5;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 4, , 5]);
                        return [4 /*yield*/, this.getTeamForCurrentUser()];
                    case 1:
                        currentTeam = _a.sent();
                        email_1 = document.getElementById("email").value;
                        return [4 /*yield*/, this.getListEmployees()];
                    case 2:
                        employeesEmails = _a.sent();
                        existingItems = employeesEmails.filter(function (item) { return item === email_1; });
                        if (existingItems.length > 0) {
                            alert("Item with email ".concat(email_1, " already exists in another team."));
                            return [2 /*return*/];
                        }
                        return [4 /*yield*/, sp.web.lists.getByTitle("Employees").items.add({
                                Fullname: document.getElementById("fullName").value,
                                Email: email_1,
                                Role: document.getElementById("role").value,
                                Team: currentTeam,
                            })];
                    case 3:
                        addItem = _a.sent();
                        void this.getAllItems();
                        alert("Item created successfully with ID: ".concat(addItem.data.ID));
                        return [3 /*break*/, 5];
                    case 4:
                        e_5 = _a.sent();
                        console.error(e_5);
                        return [3 /*break*/, 5];
                    case 5: return [2 /*return*/];
                }
            });
        }); };
        _this.state = {
            selectedButton: null, // State to manage popup visibility
        };
        return _this;
    }
    TeamOverview.prototype.componentDidMount = function () {
        // eslint-disable-next-line @typescript-eslint/no-floating-promises
        this.getAllItems(); // Call the function when the component is mounted
    };
    TeamOverview.prototype.render = function () {
        var selectedButton = this.state.selectedButton;
        return (React.createElement("div", { className: styles.teamOverview },
            React.createElement("section", null,
                React.createElement("h1", { className: styles.title }, "TeamView: Teamwork divides the task and multiplies the success!"),
                React.createElement("div", { id: "allItems" }),
                React.createElement("div", { id: "task-details-popup" }),
                React.createElement("h1", { className: styles.title }, "....")),
            React.createElement("section", null,
                React.createElement("div", { className: styles.buttonSection },
                    React.createElement("div", { className: styles.button },
                        React.createElement("span", { className: styles.label, onClick: this.handleButtonClick }, "Add Member \uD83E\uDD35"))),
                selectedButton && (React.createElement("div", { className: styles.teamOv },
                    React.createElement("form", { onSubmit: function (event) { return event.preventDefault(); } },
                        React.createElement("div", { className: styles.itemField },
                            React.createElement("div", { className: styles.fieldLabel }, "Full Name"),
                            React.createElement("input", { className: styles.fieldInput, type: "text", id: "fullName" })),
                        React.createElement("div", { className: styles.itemField },
                            React.createElement("div", { className: styles.fieldLabel }, "Email"),
                            React.createElement("input", { className: styles.fieldInput, type: "text", id: "email" })),
                        React.createElement("div", { className: styles.itemField },
                            React.createElement("div", { className: styles.fieldLabel }, "Role"),
                            React.createElement("select", { className: styles.selectField, id: "role" },
                                React.createElement("option", { value: "Agent" }, "Agent"),
                                React.createElement("option", { value: "Team Lead" }, "Team Lead"))),
                        React.createElement("div", { className: styles.buttonSection },
                            React.createElement("div", { className: styles.button },
                                React.createElement("span", { className: styles.label, onClick: this.createItem }, "Create")),
                            React.createElement("div", { className: styles.button },
                                React.createElement("span", { className: styles.label, onClick: this.handleButtonClose }, "Cancel")))))))));
    };
    return TeamOverview;
}(React.Component));
export default TeamOverview;
//# sourceMappingURL=TeamOverview.js.map