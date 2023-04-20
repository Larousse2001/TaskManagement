import * as React from 'react';
import { ICalendarOverviewProps } from "./ICalendarOverviewProps";
import "react-big-calendar/lib/css/react-big-calendar.css";
export default class CalendarOverview extends React.Component<ICalendarOverviewProps, {
    selectedTask: any | null;
}> {
    constructor(props: ICalendarOverviewProps);
    handleTaskClick: (event: any) => void;
    resizeTask: (event: any) => void;
    moveTask: (event: any) => void;
    render(): JSX.Element;
}
//# sourceMappingURL=CalendarOverview.d.ts.map