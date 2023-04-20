define("7500e25f-ba8e-43cd-bace-809c9a6b3056_0.0.1", ["@microsoft/sp-property-pane","@microsoft/sp-core-library","@microsoft/sp-webpart-base","react","react-dom","TeamOverviewWebPartStrings","@microsoft/sp-http"], function(__WEBPACK_EXTERNAL_MODULE__26ea__, __WEBPACK_EXTERNAL_MODULE_UWqr__, __WEBPACK_EXTERNAL_MODULE_br4S__, __WEBPACK_EXTERNAL_MODULE_cDcd__, __WEBPACK_EXTERNAL_MODULE_faye__, __WEBPACK_EXTERNAL_MODULE_n2sZ__, __WEBPACK_EXTERNAL_MODULE_vlQI__) { return /******/ (function(modules) { // webpackBootstrap
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
/******/ 	return __webpack_require__(__webpack_require__.s = "w561");
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

/***/ "g9++":
/*!**************************************************************!*\
  !*** ./lib/webparts/teamOverview/components/TeamOverview.js ***!
  \**************************************************************/
/*! exports provided: default */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ "cDcd");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @microsoft/sp-http */ "vlQI");
/* harmony import */ var _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_1___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_http__WEBPACK_IMPORTED_MODULE_1__);
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

// import { SPHttpClientConfiguration } from "@microsoft/sp-http";

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
            .get("".concat(this.props.listUrl, "/_api/web/lists/getbytitle('Employees')/items?$filter='Team Name' eq 'CWS CC FIT'"), _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_1__["SPHttpClient"].configurations.v1)
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
            .post("".concat(this.props.listUrl, "/_api/web/lists/getbytitle('Employees')/items"), _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_1__["SPHttpClient"].configurations.v1, spHttpClientOptions)
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
            .post("".concat(this.props.listUrl, "/_api/web/lists/getbytitle('Employees')/items(").concat(employeeId, ")"), _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_1__["SPHttpClient"].configurations.v1, spHttpClientOptions)
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
        var element = react__WEBPACK_IMPORTED_MODULE_0__["createElement"](TeamOverview, {
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
        return (react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", null,
            react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("h1", null, "Employees"),
            react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("ul", null, employees.map(function (employee) { return (react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("li", { key: employee.Id },
                react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("img", { src: employee.PictureUrl, alt: employee.Title }),
                react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", null,
                    react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("span", null, employee.Title),
                    react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("span", null, employee.JobTitle)))); }))));
    };
    return TeamOverview;
}(react__WEBPACK_IMPORTED_MODULE_0__["Component"]));
/* harmony default export */ __webpack_exports__["default"] = (TeamOverview);


/***/ }),

/***/ "n2sZ":
/*!*********************************************!*\
  !*** external "TeamOverviewWebPartStrings" ***!
  \*********************************************/
/*! no static exports found */
/***/ (function(module, exports) {

module.exports = __WEBPACK_EXTERNAL_MODULE_n2sZ__;

/***/ }),

/***/ "vlQI":
/*!*************************************!*\
  !*** external "@microsoft/sp-http" ***!
  \*************************************/
/*! no static exports found */
/***/ (function(module, exports) {

module.exports = __WEBPACK_EXTERNAL_MODULE_vlQI__;

/***/ }),

/***/ "w561":
/*!**********************************************************!*\
  !*** ./lib/webparts/teamOverview/TeamOverviewWebPart.js ***!
  \**********************************************************/
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
/* harmony import */ var TeamOverviewWebPartStrings__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! TeamOverviewWebPartStrings */ "n2sZ");
/* harmony import */ var TeamOverviewWebPartStrings__WEBPACK_IMPORTED_MODULE_5___default = /*#__PURE__*/__webpack_require__.n(TeamOverviewWebPartStrings__WEBPACK_IMPORTED_MODULE_5__);
/* harmony import */ var _components_TeamOverview__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ./components/TeamOverview */ "g9++");
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







var TeamOverviewWebPart = /** @class */ (function (_super) {
    __extends(TeamOverviewWebPart, _super);
    function TeamOverviewWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this._isDarkTheme = false;
        _this._environmentMessage = '';
        return _this;
    }
    TeamOverviewWebPart.prototype.render = function () {
        var element = react__WEBPACK_IMPORTED_MODULE_0__["createElement"](_components_TeamOverview__WEBPACK_IMPORTED_MODULE_6__["default"], {
            description: this.properties.description,
            listUrl: this.context.pageContext.web.absoluteUrl,
            spHttpClient: this.context.spHttpClient,
            isDarkTheme: this._isDarkTheme,
            environmentMessage: this._environmentMessage,
            hasTeamsContext: !!this.context.sdks.microsoftTeams,
            userDisplayName: this.context.pageContext.user.displayName
        });
        react_dom__WEBPACK_IMPORTED_MODULE_1__["render"](element, this.domElement);
    };
    TeamOverviewWebPart.prototype.onInit = function () {
        var _this = this;
        return this._getEnvironmentMessage().then(function (message) {
            _this._environmentMessage = message;
        });
    };
    TeamOverviewWebPart.prototype._getEnvironmentMessage = function () {
        var _this = this;
        if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
            return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
                .then(function (context) {
                var environmentMessage = '';
                switch (context.app.host.name) {
                    case 'Office': // running in Office
                        environmentMessage = _this.context.isServedFromLocalhost ? TeamOverviewWebPartStrings__WEBPACK_IMPORTED_MODULE_5__["AppLocalEnvironmentOffice"] : TeamOverviewWebPartStrings__WEBPACK_IMPORTED_MODULE_5__["AppOfficeEnvironment"];
                        break;
                    case 'Outlook': // running in Outlook
                        environmentMessage = _this.context.isServedFromLocalhost ? TeamOverviewWebPartStrings__WEBPACK_IMPORTED_MODULE_5__["AppLocalEnvironmentOutlook"] : TeamOverviewWebPartStrings__WEBPACK_IMPORTED_MODULE_5__["AppOutlookEnvironment"];
                        break;
                    case 'Teams': // running in Teams
                        environmentMessage = _this.context.isServedFromLocalhost ? TeamOverviewWebPartStrings__WEBPACK_IMPORTED_MODULE_5__["AppLocalEnvironmentTeams"] : TeamOverviewWebPartStrings__WEBPACK_IMPORTED_MODULE_5__["AppTeamsTabEnvironment"];
                        break;
                    default:
                        throw new Error('Unknown host');
                }
                return environmentMessage;
            });
        }
        return Promise.resolve(this.context.isServedFromLocalhost ? TeamOverviewWebPartStrings__WEBPACK_IMPORTED_MODULE_5__["AppLocalEnvironmentSharePoint"] : TeamOverviewWebPartStrings__WEBPACK_IMPORTED_MODULE_5__["AppSharePointEnvironment"]);
    };
    TeamOverviewWebPart.prototype.onThemeChanged = function (currentTheme) {
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
    TeamOverviewWebPart.prototype.onDispose = function () {
        react_dom__WEBPACK_IMPORTED_MODULE_1__["unmountComponentAtNode"](this.domElement);
    };
    Object.defineProperty(TeamOverviewWebPart.prototype, "dataVersion", {
        get: function () {
            return _microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_2__["Version"].parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    TeamOverviewWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: TeamOverviewWebPartStrings__WEBPACK_IMPORTED_MODULE_5__["PropertyPaneDescription"]
                    },
                    groups: [
                        {
                            groupName: TeamOverviewWebPartStrings__WEBPACK_IMPORTED_MODULE_5__["BasicGroupName"],
                            groupFields: [
                                Object(_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_3__["PropertyPaneTextField"])('description', {
                                    label: TeamOverviewWebPartStrings__WEBPACK_IMPORTED_MODULE_5__["DescriptionFieldLabel"]
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return TeamOverviewWebPart;
}(_microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_4__["BaseClientSideWebPart"]));
/* harmony default export */ __webpack_exports__["default"] = (TeamOverviewWebPart);


/***/ })

/******/ })});;
//# sourceMappingURL=team-overview-web-part.js.map