import * as React from "react";
import { ITeamOverviewProps } from "./ITeamOverviewProps";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
export default class TeamOverview extends React.Component<ITeamOverviewProps, {
    selectedButton: any | null;
}> {
    constructor(props: ITeamOverviewProps);
    handleButtonClick: () => void;
    handleButtonClose: () => void;
    componentDidMount(): void;
    getAllItems: () => Promise<void>;
    private displayPopup;
    private deleteItem;
    private hidePopup;
    private getListEmployees;
    render(): React.ReactElement<ITeamOverviewProps>;
    private getTeamForCurrentUser;
    private createItem;
}
//# sourceMappingURL=TeamOverview.d.ts.map