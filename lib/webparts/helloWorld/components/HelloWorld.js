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
import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
var HelloWorld = /** @class */ (function (_super) {
    __extends(HelloWorld, _super);
    function HelloWorld() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    HelloWorld.prototype.render = function () {
        var _a = this.props, isDarkTheme = _a.isDarkTheme, hasTeamsContext = _a.hasTeamsContext, userDisplayName = _a.userDisplayName;
        return (React.createElement("section", { className: "".concat(styles.helloWorld, " ").concat(hasTeamsContext ? styles.teams : '') },
            React.createElement("div", { className: styles.welcome },
                React.createElement("img", { alt: "", src: isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png'), className: styles.welcomeImage }),
                React.createElement("h2", null,
                    "Well done, ",
                    escape(userDisplayName),
                    "!")),
            React.createElement("div", null,
                React.createElement("h3", null, "Welcome to Task Management!"),
                React.createElement("p", null, "A web-based task management application that allows users to create, edit, and track tasks, as well as prioritize and categorize them by status. The application should feature intuitive drag and drop functionality, and a streamlined user interface that enables users to view tasks in multiple ways, including a calendar view, a Gantt chart, and a list view. The application should integrate with existing project management tools and be accessible via SharePoint & Teams."),
                React.createElement("h4", null, "Application features:"),
                React.createElement("ul", { className: styles.links },
                    React.createElement("li", null,
                        React.createElement("b", null, "Dashboard/Homepage:"),
                        " The dashboard or homepage is the first screen users see after logging in. It could display a summary of their tasks, upcoming deadlines, and other relevant information. This interface could also have links to different views of their tasks."),
                    React.createElement("li", null,
                        React.createElement("b", null, "Task Creation:"),
                        " The task creation interface would allow users to create new tasks by entering a title, description, due date, priority, and status. Users could also assign the task to themselves or another user."),
                    React.createElement("li", null,
                        React.createElement("b", null, "Task List View:"),
                        " The task list view displays a list of all tasks, sorted by due date or priority. Users could filter or sort the tasks based on their preferences, and they could also perform bulk actions like editing or deleting multiple tasks."),
                    React.createElement("li", null,
                        React.createElement("b", null, "Task Detail View:"),
                        " The task detail view shows all the details of a specific task, including its title, description, due date, priority, and status. Users could edit or delete the task from this interface."),
                    React.createElement("li", null,
                        React.createElement("b", null, "Calendar View:"),
                        " The calendar view displays all tasks in a monthly or weekly calendar format. Users could easily view tasks that are due on a specific day or week, and they could also drag and drop tasks to different dates."),
                    React.createElement("li", null,
                        React.createElement("b", null, "Gantt Chart View:"),
                        " The Gantt chart view shows a visual representation of all tasks and their timelines. This interface allows users to easily see how tasks are progressing and identify potential scheduling conflicts."),
                    React.createElement("li", null,
                        React.createElement("b", null, "User Management:"),
                        " The user management interface allows administrators to add, remove, or edit user accounts. It could also enable administrators to assign different roles and permissions to users.")))));
    };
    return HelloWorld;
}(React.Component));
export default HelloWorld;
//# sourceMappingURL=HelloWorld.js.map