import * as React from "react";
import { IHelloWorldProps } from "./IHelloWorldProps";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
export default class HelloWorld extends React.Component<IHelloWorldProps, {
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
}> {
    constructor(props: IHelloWorldProps);
    componentDidMount(): Promise<void>;
    loadData: () => Promise<void>;
    private getTeamForCurrentUser;
    private getListEmployees;
    private getFilteredItemsByPriority;
    private getFilteredItemsByStatus;
    private getFilteredItemsByUser;
    render(): React.ReactElement<IHelloWorldProps>;
}
//# sourceMappingURL=HelloWorld.d.ts.map