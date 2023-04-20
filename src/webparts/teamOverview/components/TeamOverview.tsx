import * as React from 'react';
import styles from './TeamOverview.module.scss';
import { ITeamOverviewProps } from './ITeamOverviewProps';

// import { SPHttpClientConfiguration } from "@microsoft/sp-http";
import {
  SPHttpClient,
  SPHttpClientResponse
} from "@microsoft/sp-http";
import { isDark } from 'office-ui-fabric-react';

interface IEmployee {
  Id: number;
  Title: string;
  JobTitle: string;
  PictureUrl: string;
}

export interface ITeamOverviewState {
  employees: IEmployee[];
}

export default class TeamOverview extends React.Component<
  ITeamOverviewProps,
  ITeamOverviewState
> {
  constructor(props: ITeamOverviewProps) {
    super(props);
    this.state = {
      employees: [],
    };
    const {
      isDarkTheme,
      hasTeamsContext,
      userDisplayName
    } = this.props;
  }

  public componentDidMount(): void {
    this.loadEmployees();
  }

  private loadEmployees(): void {
    this.props.spHttpClient
      .get(
        `${this.props.listUrl}/_api/web/lists/getbytitle('Employees')/items?$filter='Team Name' eq 'CWS CC FIT'`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((data: any) => {
        const employees: IEmployee[] = data.value.map((item: any) => {
          return {
            Id: item.Id,
            Title: item.Title,
            JobTitle: item.JobTitle,
            PictureUrl: item.PictureUrl,
          };
        });
        this.setState({ employees });
      })
      .catch((error: any) => {
        console.log(`Error: ${error}`);
      });
  }

  private addEmployee(employee: IEmployee): void {
    const { Title, JobTitle, PictureUrl } = employee;
    const spHttpClientOptions: any = {
      body: JSON.stringify({
        Title,
        JobTitle,
        PictureUrl,
        TeamName: "CWS CC FIT", // Replace with your team name
      }),
    };

    this.props.spHttpClient
      .post(
        `${this.props.listUrl}/_api/web/lists/getbytitle('Employees')/items`,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      )
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          // Employee added successfully
          // You can perform additional actions after adding a new member
          this.loadEmployees(); // Refresh the employee list after adding a new member
        } else {
          console.log(`Error adding employee: ${response.statusText}`);
        }
      })
      .catch((error: any) => {
        console.log(`Error: ${error}`);
      });
  }

  private deleteEmployee(employeeId: number): void {
    const spHttpClientOptions: any = {
      headers: {
        "X-HTTP-Method": "DELETE",
        "IF-MATCH": "*",
      },
    };

    this.props.spHttpClient
      .post(
        `${this.props.listUrl}/_api/web/lists/getbytitle('Employees')/items(${employeeId})`,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      )
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          // Employee deleted successfully
          // You can perform additional actions after deleting a member
          this.loadEmployees(); // Refresh the employee list after deleting a member
        } else {
          console.log(`Error deleting employee: ${response.statusText}`);
        }
      })
      .catch((error: any) => {
        console.log(`Error: ${error}`);
      });
  }

  public render(): React.ReactElement<ITeamOverviewProps> {
    const element: React.ReactElement<ITeamOverviewProps> = React.createElement(
      TeamOverview,
      {
        listUrl: this.context.pageContext.web.absoluteUrl,
        spHttpClient: this.context.spHttpClient,
        description: "Your description here", // Make sure to pass the 'description' property here
        isDarkTheme: true,
        environmentMessage: "Your environment here",
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );
    
    // Construct the URL for retrieving data from the SharePoint list
    const listApiUrl = `${this.props.listUrl}/_api/web/lists/getbytitle('Employees')/items`;

    const { employees } = this.state;

    return (
      <div>
        <h1>Employees</h1>
        <ul>
          {employees.map((employee: IEmployee) => (
            <li key={employee.Id}>
              <img src={employee.PictureUrl} alt={employee.Title} />
              <div>
                <span>{employee.Title}</span>
                <span>{employee.JobTitle}</span>
              </div>
            </li>
          ))}
        </ul>
      </div>
    );
  }
}