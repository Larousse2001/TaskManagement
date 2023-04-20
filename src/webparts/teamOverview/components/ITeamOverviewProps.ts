import { SPHttpClient } from "@microsoft/sp-http";

export interface ITeamOverviewProps {
  description: string;
  listUrl: string; // Add the listUrl property
  spHttpClient: SPHttpClient; // Add the spHttpClient property
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}

// export interface ITeamOverviewProps {
//  description: string;
//  isDarkTheme: boolean;
//  environmentMessage: string;
//  hasTeamsContext: boolean;
//  userDisplayName: string;
// }
