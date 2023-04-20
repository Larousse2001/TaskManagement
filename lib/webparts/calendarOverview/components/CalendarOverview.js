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
import * as React from 'react';
import styles from "./CalendarOverview.module.scss";
import { escape } from "@microsoft/sp-lodash-subset";
import { Calendar, momentLocalizer } from "react-big-calendar";
import moment from "moment";
import "react-big-calendar/lib/css/react-big-calendar.css";
import withDragAndDrop from "react-big-calendar/lib/addons/dragAndDrop";
// import 'react-big-calendar/lib/addons/dragAndDrop/styles.scss';
var localizer = momentLocalizer(moment);
var DnDCalendar = withDragAndDrop(Calendar);
var tasks = [
    {
        id: 1,
        title: "Task 1",
        start: new Date(2023, 3, 10, 9, 0),
        end: new Date(2023, 3, 10, 12, 0),
    },
    {
        id: 2,
        title: "Task 2",
        start: new Date(2023, 3, 11, 14, 0),
        end: new Date(2023, 3, 11, 17, 0),
    },
];
var CalendarOverview = /** @class */ (function (_super) {
    __extends(CalendarOverview, _super);
    function CalendarOverview(props) {
        var _this = _super.call(this, props) || this;
        _this.handleTaskClick = function (event) {
            _this.setState({
                selectedTask: event,
            });
        };
        _this.resizeTask = function (event) {
            // Update the selectedTask property in the state with the updated event
            _this.setState(function (prevState) { return ({
                selectedTask: __assign(__assign({}, prevState.selectedTask), { start: event.start, end: event.end })
            }); });
        };
        _this.moveTask = function (event) {
            // Update the selectedTask property in the state with the updated event
            _this.setState(function (prevState) { return ({
                selectedTask: __assign(__assign({}, prevState.selectedTask), { start: event.start, end: event.end })
            }); });
        };
        _this.state = {
            selectedTask: null,
        };
        return _this;
    }
    CalendarOverview.prototype.render = function () {
        var _this = this;
        var _a = this.props, hasTeamsContext = _a.hasTeamsContext, userDisplayName = _a.userDisplayName;
        var selectedTask = this.state.selectedTask;
        return (React.createElement("section", { className: "".concat(styles.calendarOverview, " ").concat(hasTeamsContext ? styles.teams : "") },
            React.createElement("div", { className: styles.welcome },
                React.createElement("h2", null,
                    "Well done, ",
                    escape(userDisplayName),
                    "!")),
            React.createElement("div", null,
                React.createElement("h3", null, "Welcome Gantt Chart !"),
                React.createElement(DnDCalendar, { localizer: localizer, defaultDate: moment().toDate(), startAccessor: "start", endAccessor: "end", events: tasks, defaultView: "month", views: ["day", "week", "month", "agenda", "work_week"], style: { height: 500 }, 
                    // onSelectEvent={this.handleTaskClick}
                    onDoubleClickEvent: this.handleTaskClick, onEventDrop: this.moveTask, onEventResize: this.resizeTask, popup: true, resizable: true })),
            selectedTask && (React.createElement("div", { className: styles["task-details-popup"] },
                React.createElement("h2", null, "Task Details"),
                React.createElement("p", null,
                    React.createElement("b", null, "ID: "),
                    selectedTask.id),
                React.createElement("p", null,
                    React.createElement("b", null, "Title: "),
                    selectedTask.title),
                React.createElement("p", null,
                    React.createElement("b", null, "Start: "),
                    selectedTask.start.toString()),
                React.createElement("p", null,
                    React.createElement("b", null, "End: "),
                    selectedTask.end.toString()),
                React.createElement("button", { onClick: function () { return _this.setState({ selectedTask: null }); } }, "Close"),
                React.createElement("button", { onClick: function () { return _this.setState({ selectedTask: null }); } }, "Update"),
                React.createElement("button", { onClick: function () { return _this.setState({ selectedTask: null }); } }, "Delete")))));
    };
    return CalendarOverview;
}(React.Component));
export default CalendarOverview;
//# sourceMappingURL=CalendarOverview.js.map