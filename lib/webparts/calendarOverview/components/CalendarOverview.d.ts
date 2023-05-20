import * as React from "react";
import { ICalendarOverviewProps } from "./ICalendarOverviewProps";
import "react-big-calendar/lib/css/react-big-calendar.css";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
export interface Event {
    id: number;
    title: string;
    start: Date;
    end: Date;
    priority: string;
    status: string;
    assignedto: string;
}
export default class CalendarOverview extends React.Component<ICalendarOverviewProps, {
    selectedTask: any | null;
    selectedButton: any | null;
    showLegend: any | null;
    events: Event[];
    eventsLoaded: true | false;
    options: any[];
    selectedValue: any[];
}> {
    constructor(props: ICalendarOverviewProps);
    componentDidMount(): Promise<void>;
    handleTaskClick: (event: any) => void;
    moveTask: (event: any) => void;
    handleButtonClick: () => void;
    loadEvents: () => Promise<void>;
    private getTeamForCurrentUser;
    private getListEmployees;
    private getFilteredItemsByStatus;
    private getFilteredItemsByPriority;
    private getTasks;
    private formatDate;
    private selectOptions;
    private handleStatusChange;
    private handlePriorityChange;
    private readItem;
    render(): JSX.Element;
    private createItem;
    private getItem;
    private updateItem;
    private deleteItem;
}
//# sourceMappingURL=CalendarOverview.d.ts.map