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
/* eslint-disable @typescript-eslint/no-empty-function */
/* eslint-disable no-void */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable react/self-closing-comp */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @rushstack/no-new-null */
import * as React from "react";
import styles from "./CalendarOverview.module.scss";
import { Calendar, momentLocalizer } from "react-big-calendar";
import "react-big-calendar/lib/css/react-big-calendar.css";
import moment from "moment";
import withDragAndDrop from "react-big-calendar/lib/addons/dragAndDrop";
// import 'react-big-calendar/lib/addons/dragAndDrop/styles.scss';
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import Multiselect from "multiselect-react-dropdown";
var localizer = momentLocalizer(moment);
// eslint-disable-next-line @typescript-eslint/no-explicit-any
var DnDCalendar = withDragAndDrop(Calendar);
var x = "x"; // priority variable
var y = "y"; // status variable
var CalendarOverview = /** @class */ (function (_super) {
    __extends(CalendarOverview, _super);
    function CalendarOverview(props) {
        var _this = _super.call(this, props) || this;
        _this.handleTaskClick = function (event) {
            _this.setState({
                selectedTask: event,
            });
        };
        _this.moveTask = function (event) {
            _this.setState(function (prevState) { return ({
                selectedTask: __assign(__assign({}, prevState.selectedTask), { start: event.start, end: event.end }),
            }); });
        };
        _this.handleButtonClick = function () {
            _this.setState({
                selectedButton: true, // Open popup when button is clicked
            });
        };
        _this.loadEvents = function () { return __awaiter(_this, void 0, void 0, function () {
            var currentTeam, items, events, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        return [4 /*yield*/, this.getTeamForCurrentUser()];
                    case 1:
                        currentTeam = _a.sent();
                        return [4 /*yield*/, sp.web.lists
                                .getByTitle("Tasks")
                                .items.filter("Team eq '".concat(currentTeam, "'"))
                                .get()];
                    case 2:
                        items = _a.sent();
                        events = items.map(function (item) { return ({
                            id: item.ID,
                            title: item.Title,
                            start: new Date(item.StartDate0),
                            end: new Date(item.EndDate0),
                            priority: item.Priority,
                            status: item.Status,
                            assignedto: item.AssignedTo0,
                        }); });
                        this.setState({ events: events, eventsLoaded: true });
                        return [3 /*break*/, 4];
                    case 3:
                        error_1 = _a.sent();
                        console.error(error_1);
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
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
        // Get filtered items by status
        _this.getFilteredItemsByStatus = function (status) { return __awaiter(_this, void 0, void 0, function () {
            var currentTeam, items, events_1, currentTeam, items, events_2, currentTeam, items, events_3, e_3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 12, , 13]);
                        if (!(status === "all")) return [3 /*break*/, 1];
                        void this.getTasks();
                        return [3 /*break*/, 11];
                    case 1:
                        if (!(x !== "x" || x !== "x")) return [3 /*break*/, 8];
                        if (!(x === "importanturg")) return [3 /*break*/, 4];
                        return [4 /*yield*/, this.getTeamForCurrentUser()];
                    case 2:
                        currentTeam = _a.sent();
                        return [4 /*yield*/, sp.web.lists
                                .getByTitle("Tasks")
                                .items.filter("Team eq '".concat(currentTeam, "' and (Priority eq 'Urgent' or Priority eq 'Important' or Priority eq 'Important and urgent') and Status eq '").concat(status, "'"))
                                .get()];
                    case 3:
                        items = _a.sent();
                        events_1 = [];
                        items.forEach(function (item) {
                            var task = {
                                id: item.ID,
                                title: item.Title,
                                start: new Date(item.StartDate0),
                                end: new Date(item.EndDate0),
                                priority: item.Priority,
                                status: item.Status,
                                assignedto: item.AssignedTo0,
                            };
                            events_1.push(task);
                        });
                        this.setState({ events: events_1, eventsLoaded: true });
                        return [3 /*break*/, 7];
                    case 4: return [4 /*yield*/, this.getTeamForCurrentUser()];
                    case 5:
                        currentTeam = _a.sent();
                        return [4 /*yield*/, sp.web.lists
                                .getByTitle("Tasks")
                                .items.filter("Team eq '".concat(currentTeam, "' and Status eq '").concat(status, "' and Priority eq '").concat(x, "'"))
                                .get()];
                    case 6:
                        items = _a.sent();
                        events_2 = [];
                        items.forEach(function (item) {
                            var task = {
                                id: item.ID,
                                title: item.Title,
                                start: new Date(item.StartDate0),
                                end: new Date(item.EndDate0),
                                priority: item.Priority,
                                status: item.Status,
                                assignedto: item.AssignedTo0,
                            };
                            events_2.push(task);
                        });
                        this.setState({ events: events_2, eventsLoaded: true });
                        _a.label = 7;
                    case 7: return [3 /*break*/, 11];
                    case 8: return [4 /*yield*/, this.getTeamForCurrentUser()];
                    case 9:
                        currentTeam = _a.sent();
                        return [4 /*yield*/, sp.web.lists
                                .getByTitle("Tasks")
                                .items.filter("Team eq '".concat(currentTeam, "' and Status eq '").concat(status, "'"))
                                .get()];
                    case 10:
                        items = _a.sent();
                        events_3 = [];
                        items.forEach(function (item) {
                            var task = {
                                id: item.ID,
                                title: item.Title,
                                start: new Date(item.StartDate0),
                                end: new Date(item.EndDate0),
                                priority: item.Priority,
                                status: item.Status,
                                assignedto: item.AssignedTo0,
                            };
                            events_3.push(task);
                        });
                        this.setState({ events: events_3, eventsLoaded: true });
                        _a.label = 11;
                    case 11: return [3 /*break*/, 13];
                    case 12:
                        e_3 = _a.sent();
                        console.error(e_3);
                        // Return an empty array in case of any error
                        return [2 /*return*/, []];
                    case 13: return [2 /*return*/];
                }
            });
        }); };
        // Get filtered items by priority
        _this.getFilteredItemsByPriority = function (priority) { return __awaiter(_this, void 0, void 0, function () {
            var currentTeam, items, events_4, currentTeam, items, events_5, currentTeam, items, events_6, currentTeam, items, events_7, e_4;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 15, , 16]);
                        if (!(priority === "all")) return [3 /*break*/, 1];
                        void this.getTasks();
                        return [3 /*break*/, 14];
                    case 1:
                        if (!(priority === "importanturg")) return [3 /*break*/, 4];
                        return [4 /*yield*/, this.getTeamForCurrentUser()];
                    case 2:
                        currentTeam = _a.sent();
                        return [4 /*yield*/, sp.web.lists
                                .getByTitle("Tasks")
                                .items.filter("Team eq '".concat(currentTeam, "' and (Priority eq 'Urgent' or Priority eq 'Important' or Priority eq 'Important and urgent')"))
                                .get()];
                    case 3:
                        items = _a.sent();
                        events_4 = [];
                        items.forEach(function (item) {
                            var task = {
                                id: item.ID,
                                title: item.Title,
                                start: new Date(item.StartDate0),
                                end: new Date(item.EndDate0),
                                priority: item.Priority,
                                status: item.Status,
                                assignedto: item.AssignedTo0,
                            };
                            events_4.push(task);
                        });
                        this.setState({ events: events_4, eventsLoaded: true });
                        return [3 /*break*/, 14];
                    case 4:
                        if (!(y !== "y")) return [3 /*break*/, 11];
                        if (!(y !== "all")) return [3 /*break*/, 7];
                        return [4 /*yield*/, this.getTeamForCurrentUser()];
                    case 5:
                        currentTeam = _a.sent();
                        return [4 /*yield*/, sp.web.lists
                                .getByTitle("Tasks")
                                .items.filter("Team eq '".concat(currentTeam, "' and Priority eq '").concat(priority, "'"))
                                .get()];
                    case 6:
                        items = _a.sent();
                        events_5 = [];
                        items.forEach(function (item) {
                            var task = {
                                id: item.ID,
                                title: item.Title,
                                start: new Date(item.StartDate0),
                                end: new Date(item.EndDate0),
                                priority: item.Priority,
                                status: item.Status,
                                assignedto: item.AssignedTo0,
                            };
                            events_5.push(task);
                        });
                        this.setState({ events: events_5, eventsLoaded: true });
                        return [3 /*break*/, 10];
                    case 7: return [4 /*yield*/, this.getTeamForCurrentUser()];
                    case 8:
                        currentTeam = _a.sent();
                        return [4 /*yield*/, sp.web.lists
                                .getByTitle("Tasks")
                                .items.filter("Team eq '".concat(currentTeam, "' and Priority eq '").concat(priority, "' and Status eq '").concat(y, "'"))
                                .get()];
                    case 9:
                        items = _a.sent();
                        events_6 = [];
                        items.forEach(function (item) {
                            var task = {
                                id: item.ID,
                                title: item.Title,
                                start: new Date(item.StartDate0),
                                end: new Date(item.EndDate0),
                                priority: item.Priority,
                                status: item.Status,
                                assignedto: item.AssignedTo0,
                            };
                            events_6.push(task);
                        });
                        this.setState({ events: events_6, eventsLoaded: true });
                        _a.label = 10;
                    case 10: return [3 /*break*/, 14];
                    case 11: return [4 /*yield*/, this.getTeamForCurrentUser()];
                    case 12:
                        currentTeam = _a.sent();
                        return [4 /*yield*/, sp.web.lists
                                .getByTitle("Tasks")
                                .items.filter("Team eq '".concat(currentTeam, "' and Priority eq '").concat(priority, "'"))
                                .get()];
                    case 13:
                        items = _a.sent();
                        events_7 = [];
                        items.forEach(function (item) {
                            var task = {
                                id: item.ID,
                                title: item.Title,
                                start: new Date(item.StartDate0),
                                end: new Date(item.EndDate0),
                                priority: item.Priority,
                                status: item.Status,
                                assignedto: item.AssignedTo0,
                            };
                            events_7.push(task);
                        });
                        this.setState({ events: events_7, eventsLoaded: true });
                        _a.label = 14;
                    case 14: return [3 /*break*/, 16];
                    case 15:
                        e_4 = _a.sent();
                        console.error(e_4);
                        // Return an empty array in case of any error
                        return [2 /*return*/, []];
                    case 16: return [2 /*return*/];
                }
            });
        }); };
        // Get Tasks from SharePoint
        _this.getTasks = function () { return __awaiter(_this, void 0, void 0, function () {
            var currentTeam, items, events_8, e_5;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        return [4 /*yield*/, this.getTeamForCurrentUser()];
                    case 1:
                        currentTeam = _a.sent();
                        return [4 /*yield*/, sp.web.lists
                                .getByTitle("Tasks")
                                .items.filter("Team eq '".concat(currentTeam, "'"))
                                .get()];
                    case 2:
                        items = _a.sent();
                        events_8 = [];
                        items.forEach(function (item) {
                            var task = {
                                id: item.ID,
                                title: item.Title,
                                start: new Date(item.StartDate0),
                                end: new Date(item.EndDate0),
                                priority: item.Priority,
                                status: item.Status,
                                assignedto: item.AssignedTo0,
                            };
                            var isDuplicate = false;
                            for (var i = 0; i < events_8.length; i++) {
                                if (JSON.stringify(events_8[i]) === JSON.stringify(task)) {
                                    isDuplicate = true;
                                    break;
                                }
                            }
                            if (!isDuplicate) {
                                events_8.push(task);
                            }
                        });
                        this.setState({ events: events_8, eventsLoaded: true });
                        return [3 /*break*/, 4];
                    case 3:
                        e_5 = _a.sent();
                        console.error(e_5);
                        return [2 /*return*/, []];
                    case 4: return [2 /*return*/];
                }
            });
        }); };
        _this.formatDate = function (date) {
            var isoDate = new Date(date).toISOString().substring(0, 10);
            var dateParts = isoDate.split("-");
            return "".concat(dateParts[0], "-").concat(dateParts[1], "-").concat(dateParts[2]);
        };
        // Select Options List
        _this.selectOptions = function () { return __awaiter(_this, void 0, void 0, function () {
            var employeesFullNames, options, index, error_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this.getListEmployees()];
                    case 1:
                        employeesFullNames = _a.sent();
                        options = [];
                        for (index = 0; index < employeesFullNames.length; index++) {
                            options.push({ key: employeesFullNames[index], id: index });
                        }
                        return [2 /*return*/, options];
                    case 2:
                        error_2 = _a.sent();
                        console.error(error_2);
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        }); };
        _this.handleStatusChange = function (event) { return __awaiter(_this, void 0, void 0, function () {
            var status;
            return __generator(this, function (_a) {
                status = event.target.value;
                y = event.target.value;
                // eslint-disable-next-line no-void
                void this.getFilteredItemsByStatus(status);
                return [2 /*return*/];
            });
        }); };
        _this.handlePriorityChange = function (event) { return __awaiter(_this, void 0, void 0, function () {
            var priority;
            return __generator(this, function (_a) {
                priority = event.target.value;
                console.log(x);
                x = event.target.value;
                console.log(x);
                // eslint-disable-next-line no-void
                void this.getFilteredItemsByPriority(priority);
                return [2 /*return*/];
            });
        }); };
        _this.readItem = function () { return __awaiter(_this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                this.handleButtonClick();
                // eslint-disable-next-line no-void
                // eslint-disable-next-line @typescript-eslint/no-floating-promises
                this.getItem();
                return [2 /*return*/];
            });
        }); };
        //Create Item
        _this.createItem = function () { return __awaiter(_this, void 0, void 0, function () {
            var currentTeam, startDate, endDate, addItem, e_6;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 4, , 5]);
                        return [4 /*yield*/, this.getTeamForCurrentUser()];
                    case 1:
                        currentTeam = _a.sent();
                        startDate = document.getElementById("startdate").value;
                        endDate = document.getElementById("enddate")
                            .value;
                        if (endDate < startDate) {
                            alert("End date must be greater than start date.");
                            return [2 /*return*/];
                        }
                        return [4 /*yield*/, sp.web.lists.getByTitle("Tasks").items.add({
                                Title: document.getElementById("title").value,
                                AssignedTo0: this.state.selectedValue
                                    .map(function (item) { return item.key; })
                                    .join(", "),
                                StartDate0: document.getElementById("startdate")
                                    .value,
                                EndDate0: document.getElementById("enddate")
                                    .value,
                                Status: document.getElementById("stat").value,
                                Priority: document.getElementById("pri").value,
                                Team: currentTeam,
                            })];
                    case 2:
                        addItem = _a.sent();
                        alert("Item created successfully with ID: ".concat(addItem.data.ID));
                        void this.getTasks();
                        return [4 /*yield*/, this.loadEvents()];
                    case 3:
                        _a.sent();
                        this.setState({ selectedButton: null });
                        return [3 /*break*/, 5];
                    case 4:
                        e_6 = _a.sent();
                        console.error(e_6);
                        return [3 /*break*/, 5];
                    case 5: return [2 /*return*/];
                }
            });
        }); };
        //Get Item by ID
        _this.getItem = function () { return __awaiter(_this, void 0, void 0, function () {
            var id, item, e_7;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 4, , 5]);
                        id = document.getElementById("itemID")
                            .value;
                        console.log(id);
                        if (!(id > 0)) return [3 /*break*/, 2];
                        return [4 /*yield*/, sp.web.lists
                                .getByTitle("Tasks")
                                .items.getById(id)
                                .get()];
                    case 1:
                        item = _a.sent();
                        document.getElementById("title").value =
                            item.Title;
                        document.getElementById("satus").value =
                            item.Status;
                        document.getElementById("priority").value =
                            item.Priority;
                        document.getElementById("startdate").value =
                            this.formatDate(item.StartDate0);
                        document.getElementById("enddate").value =
                            this.formatDate(item.EndDate0);
                        document.getElementById("assignedto").value =
                            item.AssignedTo0;
                        return [3 /*break*/, 3];
                    case 2:
                        alert("Please enter a valid item id.");
                        _a.label = 3;
                    case 3: return [3 /*break*/, 5];
                    case 4:
                        e_7 = _a.sent();
                        console.error(e_7);
                        return [3 /*break*/, 5];
                    case 5: return [2 /*return*/];
                }
            });
        }); };
        //Update Item
        _this.updateItem = function () { return __awaiter(_this, void 0, void 0, function () {
            var currentTeam, startDate, endDate, id, itemUpdate, e_8;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 6, , 7]);
                        return [4 /*yield*/, this.getTeamForCurrentUser()];
                    case 1:
                        currentTeam = _a.sent();
                        startDate = document.getElementById("startdate").value;
                        endDate = document.getElementById("enddate")
                            .value;
                        if (endDate < startDate) {
                            alert("End date must be greater than start date.");
                            return [2 /*return*/];
                        }
                        id = document.getElementById("itemID")
                            .value;
                        if (!(id > 0)) return [3 /*break*/, 4];
                        return [4 /*yield*/, sp.web.lists
                                .getByTitle("Tasks")
                                .items.getById(id)
                                .update({
                                Title: document.getElementById("title").value,
                                AssignedTo0: this.state.selectedValue
                                    .map(function (item) { return item.key; })
                                    .join(", "),
                                StartDate0: document.getElementById("startdate").value,
                                EndDate0: document.getElementById("enddate")
                                    .value,
                                Status: document.getElementById("status")
                                    .value,
                                Priority: document.getElementById("priority")
                                    .value,
                                Team: currentTeam,
                            })];
                    case 2:
                        itemUpdate = _a.sent();
                        alert("Item with ID: ".concat(id, " updated successfully!"));
                        void this.getTasks();
                        return [4 /*yield*/, this.loadEvents()];
                    case 3:
                        _a.sent();
                        this.setState({ selectedTask: null });
                        return [3 /*break*/, 5];
                    case 4:
                        alert("Please enter a valid item id.");
                        _a.label = 5;
                    case 5: return [3 /*break*/, 7];
                    case 6:
                        e_8 = _a.sent();
                        console.error(e_8);
                        return [3 /*break*/, 7];
                    case 7: return [2 /*return*/];
                }
            });
        }); };
        //Delete Item
        _this.deleteItem = function () { return __awaiter(_this, void 0, void 0, function () {
            var id, deleteItem, e_9;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 4, , 5]);
                        id = parseInt(document.getElementById("itemID").value);
                        if (!(id > 0)) return [3 /*break*/, 2];
                        return [4 /*yield*/, sp.web.lists
                                .getByTitle("Tasks")
                                .items.getById(id)
                                .delete()];
                    case 1:
                        deleteItem = _a.sent();
                        console.log(deleteItem);
                        // eslint-disable-next-line @typescript-eslint/no-floating-promises
                        this.getTasks();
                        alert("Item ID: ".concat(id, " deleted successfully!"));
                        return [3 /*break*/, 3];
                    case 2:
                        alert("Please enter a valid item id.");
                        _a.label = 3;
                    case 3: return [3 /*break*/, 5];
                    case 4:
                        e_9 = _a.sent();
                        console.error(e_9);
                        return [3 /*break*/, 5];
                    case 5: return [2 /*return*/];
                }
            });
        }); };
        _this.state = {
            showLegend: null,
            selectedTask: null,
            selectedButton: null,
            events: [],
            eventsLoaded: false,
            options: [],
            selectedValue: [],
        };
        return _this;
    }
    CalendarOverview.prototype.componentDidMount = function () {
        return __awaiter(this, void 0, void 0, function () {
            var options;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        // eslint-disable-next-line no-void
                        void this.getTasks();
                        return [4 /*yield*/, this.loadEvents()];
                    case 1:
                        _a.sent();
                        return [4 /*yield*/, this.selectOptions()];
                    case 2:
                        options = _a.sent();
                        this.setState({ options: options });
                        return [2 /*return*/];
                }
            });
        });
    };
    CalendarOverview.prototype.render = function () {
        var _this = this;
        var hasTeamsContext = this.props.hasTeamsContext;
        var selectedTask = this.state.selectedTask;
        var selectedButton = this.state.selectedButton;
        var _a = this.state, eventsLoaded = _a.eventsLoaded, events = _a.events;
        var showLegend = this.state.showLegend;
        var handleLegendClick = function (event) {
            if (showLegend === true) {
                _this.setState({
                    showLegend: null,
                });
            }
            else {
                _this.setState({
                    showLegend: true,
                });
            }
        };
        return (React.createElement("section", { className: "".concat(styles.calendarOverview, " ").concat(hasTeamsContext ? styles.teams : "") },
            React.createElement("div", { className: styles.welcome },
                React.createElement("h2", null, "TaskMaster: Take Control of Your Schedule!")),
            React.createElement("br", null),
            React.createElement("div", { className: "special" },
                React.createElement("div", { className: styles.buttonSec },
                    React.createElement("div", { className: styles.itemF },
                        React.createElement("select", { className: styles.inputField, id: "status", onChange: this.handleStatusChange },
                            React.createElement("option", { value: "all" }, "All"),
                            React.createElement("option", { value: "Active" }, "Active"),
                            React.createElement("option", { value: "Completed" }, "Completed"),
                            React.createElement("option", { value: "On Hold" }, "On Hold"),
                            React.createElement("option", { value: "Cancelled" }, "Cancelled")),
                        React.createElement("div", { className: styles.fieldL },
                            React.createElement("b", null, " Filter By Status"))),
                    React.createElement("div", { className: styles.itemF },
                        React.createElement("select", { className: styles.inputField, id: "priority", onChange: this.handlePriorityChange },
                            React.createElement("option", { value: "all" }, "All"),
                            React.createElement("option", { value: "Important" }, "Important"),
                            React.createElement("option", { value: "importanturg" }, "Important and urgent"),
                            React.createElement("option", { value: "Urgent" }, "Urgent"),
                            React.createElement("option", { value: "Neither" }, "Neither")),
                        React.createElement("div", { className: styles.fieldL },
                            React.createElement("b", null, " Filter By Priority"))),
                    React.createElement("div", { className: styles.button },
                        React.createElement("span", { className: styles.label, onClick: handleLegendClick }, "Legend")))),
            showLegend && (React.createElement("div", { className: styles.legend },
                React.createElement("div", null,
                    React.createElement("span", { style: { color: "RGBA(0, 169, 235, 1)" } },
                        React.createElement("b", null, "Neither"))),
                React.createElement("div", null,
                    React.createElement("span", { style: { color: "RGBA(254, 211, 76, 1)" } },
                        React.createElement("b", null, "Important"))),
                React.createElement("div", null,
                    React.createElement("span", { style: { color: "RGBA(255, 153, 18, 1)" } },
                        React.createElement("b", null, "Urgent"))),
                React.createElement("div", null,
                    React.createElement("span", { style: { color: "RGBA(250, 0, 87, 1)" } },
                        React.createElement("b", null, "Important and urgent"))))),
            React.createElement("br", null),
            React.createElement("br", null),
            React.createElement("div", null, !eventsLoaded ? (React.createElement("div", null,
                React.createElement("b", null, "Loading..."))) : (React.createElement(DnDCalendar, { localizer: localizer, defaultDate: moment().toDate(), startAccessor: "start", endAccessor: "end", events: events, defaultView: "month", views: ["day", "week", "month", "agenda"], style: { height: 500 }, onDoubleClickEvent: this.handleTaskClick, onEventDrop: this.moveTask, popup: true, eventPropGetter: function (event) {
                    var style = {
                        backgroundColor: "RGBA(0, 169, 235, 1)",
                    };
                    if (event.priority === "Important") {
                        style.backgroundColor = "RGBA(254, 211, 76, 1)";
                    }
                    if (event.priority === "Urgent") {
                        style.backgroundColor = "RGBA(255, 153, 18, 1)";
                    }
                    if (event.priority === "Important and urgent") {
                        style.backgroundColor = "RGBA(250, 0, 87, 1)";
                    }
                    return {
                        style: style,
                    };
                } }))),
            React.createElement("br", null),
            React.createElement("div", { className: styles.buttonSection },
                React.createElement("div", { className: styles.button },
                    React.createElement("span", { className: styles.label, onClick: this.handleButtonClick }, "Add a Task \uD83D\uDCDD"))),
            selectedButton && (React.createElement("div", { className: styles.teamOv },
                React.createElement("form", { onSubmit: function (event) { return event.preventDefault(); } },
                    React.createElement("div", { className: styles.itemField },
                        React.createElement("div", { className: styles.fieldLabel }, "Title"),
                        React.createElement("input", { className: styles.fieldInput, type: "text", id: "title" })),
                    React.createElement("div", { className: styles.itemField },
                        React.createElement("div", { className: styles.fieldLabel }, "Assigned To"),
                        React.createElement(Multiselect, { id: "assignedto", options: this.state.options, selectedValues: this.state.selectedValue, onKeyPressFn: function noRefCheck() { }, onRemove: function noRefCheck() { }, onSearch: function noRefCheck() { }, onSelect: function (selectedList, selectedItem) {
                                _this.setState({ selectedValue: selectedList });
                            }, displayValue: "key" // Property name to display in the dropdown options
                            , showCheckbox: true, placeholder: "Choose the member" })),
                    React.createElement("div", { className: styles.itemField },
                        React.createElement("div", { className: styles.fieldLabel }, "Start Date"),
                        React.createElement("input", { className: styles.dateInput, type: "date", id: "startdate" })),
                    React.createElement("div", { className: styles.itemField },
                        React.createElement("div", { className: styles.fieldLabel }, "End Date"),
                        React.createElement("input", { className: styles.dateInput, type: "date", id: "enddate" })),
                    React.createElement("div", { className: styles.itemField },
                        React.createElement("div", { className: styles.fieldLabel }, "Status"),
                        React.createElement("select", { className: styles.selectField, id: "stat" },
                            React.createElement("option", { value: "Active" }, "Active"),
                            React.createElement("option", { value: "Completed" }, "Completed"),
                            React.createElement("option", { value: "On Hold" }, "On Hold"),
                            React.createElement("option", { value: "Cancelled" }, "Cancelled"))),
                    React.createElement("div", { className: styles.itemField },
                        React.createElement("div", { className: styles.fieldLabel }, "Priority"),
                        React.createElement("select", { className: styles.selectField, id: "pri" },
                            React.createElement("option", { value: "Important" }, "Important"),
                            React.createElement("option", { value: "Important and urgent" }, "Important and urgent"),
                            React.createElement("option", { value: "Urgent" }, "Urgent"),
                            React.createElement("option", { value: "Neither" }, "Neither"))),
                    React.createElement("div", { className: styles.buttonSection },
                        React.createElement("div", { className: styles.button },
                            React.createElement("span", { className: styles.label, onClick: this.createItem }, "Create")),
                        React.createElement("div", { className: styles.button },
                            React.createElement("span", { className: styles.label, onClick: this.updateItem }, "Update")),
                        React.createElement("div", { className: styles.button },
                            React.createElement("span", { className: styles.label, onClick: function () { return _this.setState({ selectedButton: null }); } }, "Cancel")))))),
            selectedTask && (React.createElement("div", { className: styles["task-details-popup"] },
                React.createElement("h2", null, "Task Details"),
                React.createElement("div", { className: styles.itemField },
                    React.createElement("div", { className: styles.fieldLabel }, "ID:"),
                    React.createElement("input", { type: "text", id: "itemID", value: selectedTask.id, disabled: true })),
                React.createElement("p", null,
                    React.createElement("b", null, "Title: "),
                    selectedTask.title),
                React.createElement("p", null,
                    React.createElement("b", null, "Start: "),
                    selectedTask.start.toString()),
                React.createElement("p", null,
                    React.createElement("b", null, "End: "),
                    selectedTask.end.toString()),
                React.createElement("p", null,
                    React.createElement("b", null, "Assigned To: "),
                    selectedTask.assignedto),
                React.createElement("p", null,
                    React.createElement("b", null, "Status: "),
                    selectedTask.status),
                React.createElement("p", null,
                    React.createElement("b", null, "Priority: "),
                    selectedTask.priority),
                React.createElement("button", { onClick: this.readItem }, "Update"),
                React.createElement("button", { onClick: this.deleteItem }, "Delete"),
                React.createElement("button", { onClick: function () { return _this.setState({ selectedTask: null }); } }, "Close")))));
    };
    return CalendarOverview;
}(React.Component));
export default CalendarOverview;
//# sourceMappingURL=CalendarOverview.js.map