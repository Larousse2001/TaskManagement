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
/* eslint-disable react/self-closing-comp */
/* eslint-disable no-unused-labels */
/* eslint-disable no-unused-expressions */
/* eslint-disable @typescript-eslint/no-var-requires */
/* eslint-disable no-void */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-unused-vars */
import * as React from "react";
import styles from "./HelloWorld.module.scss";
import { escape } from "@microsoft/sp-lodash-subset";
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { ChartControl, ChartType, } from "@pnp/spfx-controls-react/lib/ChartControl";
var HelloWorld = /** @class */ (function (_super) {
    __extends(HelloWorld, _super);
    function HelloWorld(props) {
        var _this = _super.call(this, props) || this;
        _this.loadData = function () { return __awaiter(_this, void 0, void 0, function () {
            var importantAndUrgentPromise, importantPromise, urgentPromise, neitherPromise, activePromise, completedPromise, onholdPromise, cancelPromise, userslabels, promises, usersdata, team, error_1;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 12, , 13]);
                        return [4 /*yield*/, this.getFilteredItemsByPriority("Important and urgent")];
                    case 1:
                        importantAndUrgentPromise = _a.sent();
                        return [4 /*yield*/, this.getFilteredItemsByPriority("Important")];
                    case 2:
                        importantPromise = _a.sent();
                        return [4 /*yield*/, this.getFilteredItemsByPriority("Urgent")];
                    case 3:
                        urgentPromise = _a.sent();
                        return [4 /*yield*/, this.getFilteredItemsByPriority("Neither")];
                    case 4:
                        neitherPromise = _a.sent();
                        return [4 /*yield*/, this.getFilteredItemsByStatus("Active")];
                    case 5:
                        activePromise = _a.sent();
                        return [4 /*yield*/, this.getFilteredItemsByStatus("Completed")];
                    case 6:
                        completedPromise = _a.sent();
                        return [4 /*yield*/, this.getFilteredItemsByStatus("On Hold")];
                    case 7:
                        onholdPromise = _a.sent();
                        return [4 /*yield*/, this.getFilteredItemsByStatus("Cancelled")];
                    case 8:
                        cancelPromise = _a.sent();
                        return [4 /*yield*/, this.getListEmployees()];
                    case 9:
                        userslabels = _a.sent();
                        promises = userslabels.map(function (user) { return __awaiter(_this, void 0, void 0, function () {
                            return __generator(this, function (_a) {
                                switch (_a.label) {
                                    case 0: return [4 /*yield*/, this.getFilteredItemsByUser(user)];
                                    case 1: return [2 /*return*/, _a.sent()];
                                }
                            });
                        }); });
                        return [4 /*yield*/, Promise.all(promises)];
                    case 10:
                        usersdata = _a.sent();
                        return [4 /*yield*/, this.getTeamForCurrentUser()];
                    case 11:
                        team = _a.sent();
                        this.setState({
                            importantAndUrgentPromise: importantAndUrgentPromise,
                            importantPromise: importantPromise,
                            urgentPromise: urgentPromise,
                            neitherPromise: neitherPromise,
                            activePromise: activePromise,
                            completedPromise: completedPromise,
                            onholdPromise: onholdPromise,
                            cancelPromise: cancelPromise,
                            userslabels: userslabels,
                            usersdata: usersdata,
                            team: team,
                            eventsLoaded: true,
                        });
                        return [3 /*break*/, 13];
                    case 12:
                        error_1 = _a.sent();
                        console.error(error_1);
                        return [3 /*break*/, 13];
                    case 13: return [2 /*return*/];
                }
            });
        }); };
        // Function to get the current Team variable of the current user from SharePoint
        _this.getTeamForCurrentUser = function () { return __awaiter(_this, void 0, void 0, function () {
            var currentUser, userEmail, list, items, currentTeam, e_1;
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
                        e_1 = _a.sent();
                        console.error(e_1);
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        }); };
        // Get all items and return a list filled with employees' full names or an empty list
        _this.getListEmployees = function () { return __awaiter(_this, void 0, void 0, function () {
            var currentTeam, items, employeesFullNames_1, e_2;
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
                        employeesFullNames_1 = [];
                        if (items.length > 0) {
                            // Loop through the retrieved items and extract the "Fullname" field value
                            items.forEach(function (item) {
                                // Add the "Fullname" field value to the employees' full names array
                                employeesFullNames_1.push(item.Fullname);
                            });
                        }
                        // Return the employees' full names array
                        return [2 /*return*/, employeesFullNames_1];
                    case 3:
                        e_2 = _a.sent();
                        console.error(e_2);
                        // Return an empty array in case of any error
                        return [2 /*return*/, []];
                    case 4: return [2 /*return*/];
                }
            });
        }); };
        // Get filtered items by priority
        _this.getFilteredItemsByPriority = function (priority) { return __awaiter(_this, void 0, void 0, function () {
            var currentTeam, items, prioritydata, e_3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        return [4 /*yield*/, this.getTeamForCurrentUser()];
                    case 1:
                        currentTeam = _a.sent();
                        return [4 /*yield*/, sp.web.lists
                                .getByTitle("Tasks")
                                .items.filter("Team eq '".concat(currentTeam, "' and Priority eq '").concat(priority, "'"))
                                .get()];
                    case 2:
                        items = _a.sent();
                        prioritydata = items.length;
                        return [2 /*return*/, prioritydata];
                    case 3:
                        e_3 = _a.sent();
                        console.error(e_3);
                        // Return 0 in case of any error
                        return [2 /*return*/, 0];
                    case 4: return [2 /*return*/];
                }
            });
        }); };
        // Get filtered items by status
        _this.getFilteredItemsByStatus = function (status) { return __awaiter(_this, void 0, void 0, function () {
            var currentTeam, items, statusdata, e_4;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        return [4 /*yield*/, this.getTeamForCurrentUser()];
                    case 1:
                        currentTeam = _a.sent();
                        return [4 /*yield*/, sp.web.lists
                                .getByTitle("Tasks")
                                .items.filter("Team eq '".concat(currentTeam, "' and Status eq '").concat(status, "'"))
                                .get()];
                    case 2:
                        items = _a.sent();
                        statusdata = items.length;
                        return [2 /*return*/, statusdata];
                    case 3:
                        e_4 = _a.sent();
                        console.error(e_4);
                        // Return an empty array in case of any error
                        return [2 /*return*/, 0];
                    case 4: return [2 /*return*/];
                }
            });
        }); };
        // Get filtered items by user
        _this.getFilteredItemsByUser = function (user) { return __awaiter(_this, void 0, void 0, function () {
            var currentTeam, items, userdata, e_5;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        return [4 /*yield*/, this.getTeamForCurrentUser()];
                    case 1:
                        currentTeam = _a.sent();
                        return [4 /*yield*/, sp.web.lists
                                .getByTitle("Tasks")
                                .items.filter("Team eq '".concat(currentTeam, "' and AssignedTo0 eq '").concat(user, "'"))
                                .get()];
                    case 2:
                        items = _a.sent();
                        userdata = items.length;
                        return [2 /*return*/, userdata];
                    case 3:
                        e_5 = _a.sent();
                        console.error(e_5);
                        // Return an empty array in case of any error
                        return [2 /*return*/, 0];
                    case 4: return [2 /*return*/];
                }
            });
        }); };
        _this.state = {
            importantAndUrgentPromise: 0,
            importantPromise: 0,
            urgentPromise: 0,
            neitherPromise: 0,
            activePromise: 0,
            completedPromise: 0,
            onholdPromise: 0,
            cancelPromise: 0,
            userslabels: [],
            usersdata: [],
            team: "",
            eventsLoaded: false,
        };
        return _this;
    }
    HelloWorld.prototype.componentDidMount = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.loadData()];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    HelloWorld.prototype.render = function () {
        var _a = this.props, isDarkTheme = _a.isDarkTheme, hasTeamsContext = _a.hasTeamsContext, userDisplayName = _a.userDisplayName;
        var _b = this.state, importantAndUrgentPromise = _b.importantAndUrgentPromise, importantPromise = _b.importantPromise, urgentPromise = _b.urgentPromise, neitherPromise = _b.neitherPromise, activePromise = _b.activePromise, completedPromise = _b.completedPromise, onholdPromise = _b.onholdPromise, cancelPromise = _b.cancelPromise, userslabels = _b.userslabels, usersdata = _b.usersdata, team = _b.team, eventsLoaded = _b.eventsLoaded;
        // set the data
        var dataPriority = {
            labels: ["Important and urgent", "Important", "Urgent", "Neither"],
            datasets: [
                {
                    label: "Dataset of Priority",
                    fill: false,
                    lineTension: 0,
                    data: [
                        importantAndUrgentPromise,
                        importantPromise,
                        urgentPromise,
                        neitherPromise,
                    ],
                },
            ],
        };
        var dataStatus = {
            labels: ["Active", "Completed", "On Hold", "Cancelled"],
            datasets: [
                {
                    label: "Dataset of Status",
                    fill: false,
                    lineTension: 0,
                    data: [activePromise, completedPromise, onholdPromise, cancelPromise],
                },
            ],
        };
        var dataUsers = {
            labels: userslabels,
            datasets: [
                {
                    label: "Employees of " + team,
                    fill: true,
                    lineTension: 0,
                    data: usersdata,
                },
            ],
        };
        // set the options
        var optionsPriority = {
            legend: {
                display: true,
            },
            title: {
                display: true,
                text: "Visualizing Task Prioritization with a Doughnut Chart",
            },
        };
        var optionsStatus = {
            legend: {
                display: true,
            },
            title: {
                display: true,
                text: "From To-Do to Done: A Pie Chart Depicting Task Status Changes",
            },
        };
        var optionsUsers = {
            legend: {
                display: true,
            },
            title: {
                display: true,
                text: "Employee Productivity Bar Chart for Task Management",
            },
        };
        return (React.createElement("section", { className: "".concat(styles.helloWorld, " ").concat(hasTeamsContext ? styles.teams : "") },
            React.createElement("div", { className: styles.welcome },
                React.createElement("img", { alt: "", src: isDarkTheme
                        ? require("../assets/welcome-dark.png")
                        : require("../assets/welcome-light.png"), className: styles.welcomeImage }),
                React.createElement("h2", null,
                    "Well done, ",
                    escape(userDisplayName),
                    "!")),
            React.createElement("div", null,
                React.createElement("h3", null, "Welcome to Task Management!"),
                React.createElement("p", null, "A web-based task management application that allows users to create, edit, and track tasks, as well as prioritize and categorize them by status. The application should feature intuitive drag and drop functionality, and a streamlined user interface that enables users to view tasks in multiple ways, including a calendar view, a Gantt chart, and a list view. The application should integrate with existing project management tools and be accessible via SharePoint.")),
            React.createElement("div", null, !eventsLoaded ? (React.createElement("div", null,
                React.createElement("b", null, "Loading..."))) : (React.createElement(ChartControl, { type: ChartType.Doughnut, data: dataPriority, options: optionsPriority }))),
            React.createElement("div", null, !eventsLoaded ? (React.createElement("div", null,
                React.createElement("b", null))) : (React.createElement(ChartControl, { type: ChartType.Pie, data: dataStatus, options: optionsStatus }))),
            React.createElement("div", null, !eventsLoaded ? (React.createElement("div", null,
                React.createElement("b", null))) : (React.createElement(ChartControl, { type: ChartType.Bar, data: dataUsers, options: optionsUsers })))));
    };
    return HelloWorld;
}(React.Component));
export default HelloWorld;
//# sourceMappingURL=HelloWorld.js.map