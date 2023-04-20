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
// import { SPHttpClientConfiguration } from "@microsoft/sp-http";
import { SPHttpClient } from "@microsoft/sp-http";
var TeamOverview = /** @class */ (function (_super) {
    __extends(TeamOverview, _super);
    function TeamOverview(props) {
        var _this = _super.call(this, props) || this;
        _this.state = {
            employees: [],
        };
        var _a = _this.props, isDarkTheme = _a.isDarkTheme, hasTeamsContext = _a.hasTeamsContext, userDisplayName = _a.userDisplayName;
        return _this;
    }
    TeamOverview.prototype.componentDidMount = function () {
        this.loadEmployees();
    };
    TeamOverview.prototype.loadEmployees = function () {
        var _this = this;
        this.props.spHttpClient
            .get("".concat(this.props.listUrl, "/_api/web/lists/getbytitle('Employees')/items?$filter='Team Name' eq 'CWS CC FIT'"), SPHttpClient.configurations.v1)
            .then(function (response) {
            return response.json();
        })
            .then(function (data) {
            var employees = data.value.map(function (item) {
                return {
                    Id: item.Id,
                    Title: item.Title,
                    JobTitle: item.JobTitle,
                    PictureUrl: item.PictureUrl,
                };
            });
            _this.setState({ employees: employees });
        })
            .catch(function (error) {
            console.log("Error: ".concat(error));
        });
    };
    TeamOverview.prototype.addEmployee = function (employee) {
        var _this = this;
        var Title = employee.Title, JobTitle = employee.JobTitle, PictureUrl = employee.PictureUrl;
        var spHttpClientOptions = {
            body: JSON.stringify({
                Title: Title,
                JobTitle: JobTitle,
                PictureUrl: PictureUrl,
                TeamName: "CWS CC FIT", // Replace with your team name
            }),
        };
        this.props.spHttpClient
            .post("".concat(this.props.listUrl, "/_api/web/lists/getbytitle('Employees')/items"), SPHttpClient.configurations.v1, spHttpClientOptions)
            .then(function (response) {
            if (response.ok) {
                // Employee added successfully
                // You can perform additional actions after adding a new member
                _this.loadEmployees(); // Refresh the employee list after adding a new member
            }
            else {
                console.log("Error adding employee: ".concat(response.statusText));
            }
        })
            .catch(function (error) {
            console.log("Error: ".concat(error));
        });
    };
    TeamOverview.prototype.deleteEmployee = function (employeeId) {
        var _this = this;
        var spHttpClientOptions = {
            headers: {
                "X-HTTP-Method": "DELETE",
                "IF-MATCH": "*",
            },
        };
        this.props.spHttpClient
            .post("".concat(this.props.listUrl, "/_api/web/lists/getbytitle('Employees')/items(").concat(employeeId, ")"), SPHttpClient.configurations.v1, spHttpClientOptions)
            .then(function (response) {
            if (response.ok) {
                // Employee deleted successfully
                // You can perform additional actions after deleting a member
                _this.loadEmployees(); // Refresh the employee list after deleting a member
            }
            else {
                console.log("Error deleting employee: ".concat(response.statusText));
            }
        })
            .catch(function (error) {
            console.log("Error: ".concat(error));
        });
    };
    TeamOverview.prototype.render = function () {
        var element = React.createElement(TeamOverview, {
            listUrl: this.context.pageContext.web.absoluteUrl,
            spHttpClient: this.context.spHttpClient,
            description: "Your description here",
            isDarkTheme: true,
            environmentMessage: "Your environment here",
            hasTeamsContext: !!this.context.sdks.microsoftTeams,
            userDisplayName: this.context.pageContext.user.displayName
        });
        // Construct the URL for retrieving data from the SharePoint list
        var listApiUrl = "".concat(this.props.listUrl, "/_api/web/lists/getbytitle('Employees')/items");
        var employees = this.state.employees;
        return (React.createElement("div", null,
            React.createElement("h1", null, "Employees"),
            React.createElement("ul", null, employees.map(function (employee) { return (React.createElement("li", { key: employee.Id },
                React.createElement("img", { src: employee.PictureUrl, alt: employee.Title }),
                React.createElement("div", null,
                    React.createElement("span", null, employee.Title),
                    React.createElement("span", null, employee.JobTitle)))); }))));
    };
    return TeamOverview;
}(React.Component));
export default TeamOverview;
//# sourceMappingURL=TeamOverview.js.map