import * as React from 'react';
import { ITeamOverviewProps } from './ITeamOverviewProps';
interface IEmployee {
    Id: number;
    Title: string;
    JobTitle: string;
    PictureUrl: string;
}
export interface ITeamOverviewState {
    employees: IEmployee[];
}
export default class TeamOverview extends React.Component<ITeamOverviewProps, ITeamOverviewState> {
    constructor(props: ITeamOverviewProps);
    componentDidMount(): void;
    private loadEmployees;
    private addEmployee;
    private deleteEmployee;
    render(): React.ReactElement<ITeamOverviewProps>;
}
export {};
//# sourceMappingURL=TeamOverview.d.ts.map