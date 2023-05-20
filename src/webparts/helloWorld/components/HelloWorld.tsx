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
import { IHelloWorldProps } from "./IHelloWorldProps";
import { escape } from "@microsoft/sp-lodash-subset";

import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import {
  ChartControl,
  ChartType,
} from "@pnp/spfx-controls-react/lib/ChartControl";

export default class HelloWorld extends React.Component<
  IHelloWorldProps,
  {
    importantAndUrgentPromise: number;
    importantPromise: number;
    urgentPromise: number;
    neitherPromise: number;
    activePromise: number;
    completedPromise: number;
    onholdPromise: number;
    cancelPromise: number;
    userslabels: string[];
    usersdata: number[];
    team: string;
    eventsLoaded: true | false;
  }
> {
  constructor(props: IHelloWorldProps) {
    super(props);
    this.state = {
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
  }

  public async componentDidMount(): Promise<void> {
    await this.loadData();
  }

  loadData = async () => {
    try {
      const importantAndUrgentPromise = await this.getFilteredItemsByPriority(
        "Important and urgent"
      );
      const importantPromise = await this.getFilteredItemsByPriority(
        "Important"
      );
      const urgentPromise = await this.getFilteredItemsByPriority("Urgent");
      const neitherPromise = await this.getFilteredItemsByPriority("Neither");
      const activePromise = await this.getFilteredItemsByStatus("Active");
      const completedPromise = await this.getFilteredItemsByStatus("Completed");
      const onholdPromise = await this.getFilteredItemsByStatus("On Hold");
      const cancelPromise = await this.getFilteredItemsByStatus("Cancelled");

      const userslabels = await this.getListEmployees();
      const promises = userslabels.map(async (user: string) => {
        return await this.getFilteredItemsByUser(user);
      });
      const usersdata = await Promise.all(promises);

      const team = await this.getTeamForCurrentUser();

      this.setState({
        importantAndUrgentPromise,
        importantPromise,
        urgentPromise,
        neitherPromise,
        activePromise,
        completedPromise,
        onholdPromise,
        cancelPromise,
        userslabels,
        usersdata,
        team,
        eventsLoaded: true,
      });
    } catch (error) {
      console.error(error);
    }
  };

  // Function to get the current Team variable of the current user from SharePoint
  private getTeamForCurrentUser = async () => {
    try {
      const currentUser = await sp.web.currentUser();
      const userEmail = currentUser.Email;
      const list = sp.web.lists.getByTitle("Employees");
      const items = await list.items.filter(`Email eq '${userEmail}'`).get();
      if (items.length > 0) {
        const currentTeam = items[0].Team;
        return currentTeam;
      } else {
        return "";
      }
    } catch (e) {
      console.error(e);
    }
  };

  // Get all items and return a list filled with employees' full names or an empty list
  private getListEmployees = async (): Promise<string[]> => {
    try {
      const currentTeam = await this.getTeamForCurrentUser();
      const items: any[] = await sp.web.lists
        .getByTitle("Employees")
        .items.filter(`Team eq '${currentTeam}'`)
        .get();

      // Create an array to store employees' full names
      const employeesFullNames: string[] = [];

      if (items.length > 0) {
        // Loop through the retrieved items and extract the "Fullname" field value
        items.forEach((item) => {
          // Add the "Fullname" field value to the employees' full names array
          employeesFullNames.push(item.Fullname);
        });
      }

      // Return the employees' full names array
      return employeesFullNames;
    } catch (e) {
      console.error(e);
      // Return an empty array in case of any error
      return [];
    }
  };

  // Get filtered items by priority
  private getFilteredItemsByPriority = async (priority: string) => {
    try {
      const currentTeam = await this.getTeamForCurrentUser();
      const items: any[] = await sp.web.lists
        .getByTitle("Tasks")
        .items.filter(`Team eq '${currentTeam}' and Priority eq '${priority}'`)
        .get();
      const prioritydata = items.length;
      return prioritydata;
    } catch (e) {
      console.error(e);
      // Return 0 in case of any error
      return 0;
    }
  };

  // Get filtered items by status
  private getFilteredItemsByStatus = async (status: string) => {
    try {
      const currentTeam = await this.getTeamForCurrentUser();
      const items: any[] = await sp.web.lists
        .getByTitle("Tasks")
        .items.filter(`Team eq '${currentTeam}' and Status eq '${status}'`)
        .get();
      const statusdata = items.length;
      return statusdata;
    } catch (e) {
      console.error(e);
      // Return an empty array in case of any error
      return 0;
    }
  };

  // Get filtered items by user
  private getFilteredItemsByUser = async (user: string) => {
    try {
      const currentTeam = await this.getTeamForCurrentUser();
      const items: any[] = await sp.web.lists
        .getByTitle("Tasks")
        .items.filter(`Team eq '${currentTeam}' and AssignedTo0 eq '${user}'`)
        .get();
      const userdata = items.length;
      return userdata;
    } catch (e) {
      console.error(e);
      // Return an empty array in case of any error
      return 0;
    }
  };

  public render(): React.ReactElement<IHelloWorldProps> {
    const { isDarkTheme, hasTeamsContext, userDisplayName } = this.props;
    const {
      importantAndUrgentPromise,
      importantPromise,
      urgentPromise,
      neitherPromise,
      activePromise,
      completedPromise,
      onholdPromise,
      cancelPromise,
      userslabels,
      usersdata,
      team,
      eventsLoaded,
    } = this.state;

    // set the data
    const dataPriority = {
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

    const dataStatus = {
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

    const dataUsers = {
      labels: userslabels,
      datasets: [
        {
          label: "Employees of "+team,
          fill: true,
          lineTension: 0,
          data: usersdata,
        },
      ],
    };

    // set the options
    const optionsPriority = {
      legend: {
        display: true,
      },
      title: {
        display: true,
        text: "Visualizing Task Prioritization with a Doughnut Chart",
      },
    };

    const optionsStatus = {
      legend: {
        display: true,
      },
      title: {
        display: true,
        text: "From To-Do to Done: A Pie Chart Depicting Task Status Changes",
      },
    };

    const optionsUsers = {
      legend: {
        display: true,
      },
      title: {
        display: true,
        text: "Employee Productivity Bar Chart for Task Management",
      },
    };

    return (
      <section
        className={`${styles.helloWorld} ${
          hasTeamsContext ? styles.teams : ""
        }`}
      >
        <div className={styles.welcome}>
          <img
            alt=""
            src={
              isDarkTheme
                ? require("../assets/welcome-dark.png")
                : require("../assets/welcome-light.png")
            }
            className={styles.welcomeImage}
          />
          <h2>Well done, {escape(userDisplayName)}!</h2>
        </div>
        <div>
          <h3>Welcome to Task Management!</h3>
          <p>
            A web-based task management application that allows users to create,
            edit, and track tasks, as well as prioritize and categorize them by
            status. The application should feature intuitive drag and drop
            functionality, and a streamlined user interface that enables users
            to view tasks in multiple ways, including a calendar view, a Gantt
            chart, and a list view. The application should integrate with
            existing project management tools and be accessible via SharePoint.
          </p>
        </div>
        <div>
          {!eventsLoaded ? (
            <div>
              <b>Loading...</b>
            </div>
          ) : (
            <ChartControl
              type={ChartType.Doughnut}
              data={dataPriority}
              options={optionsPriority}
            />
          )}
        </div>
        <div>
          {!eventsLoaded ? (
            <div>
              <b></b>
            </div>
          ) : (
            <ChartControl
              type={ChartType.Pie}
              data={dataStatus}
              options={optionsStatus}
            />
          )}
        </div>
        <div>
          {!eventsLoaded ? (
            <div>
              <b></b>
            </div>
          ) : (
            <ChartControl
              type={ChartType.Bar}
              data={dataUsers}
              options={optionsUsers}
            />
          )}
        </div>
      </section>
    );
  }
}
