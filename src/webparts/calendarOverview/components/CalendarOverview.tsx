import * as React from 'react';
import styles from "./CalendarOverview.module.scss";
import { ICalendarOverviewProps } from "./ICalendarOverviewProps";
import { escape } from "@microsoft/sp-lodash-subset";

import { Calendar, momentLocalizer } from "react-big-calendar";
import moment from "moment";
import "react-big-calendar/lib/css/react-big-calendar.css";

import withDragAndDrop from "react-big-calendar/lib/addons/dragAndDrop";
// import 'react-big-calendar/lib/addons/dragAndDrop/styles.scss';

const localizer = momentLocalizer(moment);
const DnDCalendar = withDragAndDrop(Calendar as any);

const tasks = [
  {
    id: 1,
    title: "Task 1",
    start: new Date(2023, 3, 10, 9, 0),
    end: new Date(2023, 3, 10, 12, 0),
  },
  {
    id: 2,
    title: "Task 2",
    start: new Date(2023, 3, 11, 14, 0),
    end: new Date(2023, 3, 11, 17, 0),
  },
];

export default class CalendarOverview extends React.Component<
  ICalendarOverviewProps,
  {
    selectedTask: any | null;
  }
> {
  constructor(props: ICalendarOverviewProps) {
    super(props);
    this.state = {
      selectedTask: null,
    };
  }

  handleTaskClick = (event: any) => {
    this.setState({
      selectedTask: event,
    });
  };

  resizeTask = (event: any) => {
    this.setState(prevState => ({
      selectedTask: {
        ...prevState.selectedTask,
        start: event.start,
        end: event.end
      }
    }));
  };
  
  moveTask = (event: any) => {
    this.setState(prevState => ({
      selectedTask: {
        ...prevState.selectedTask,
        start: event.start,
        end: event.end
      }
    }));
  };  

  render() {
    const { hasTeamsContext, userDisplayName } = this.props;

    const { selectedTask } = this.state;

    return (
      <section
        className={`${styles.calendarOverview} ${
          hasTeamsContext ? styles.teams : ""
        }`}
      >
        <div className={styles.welcome}>
          <h2>Well done, {escape(userDisplayName)}!</h2>
        </div>
        <div>
          <h3>Welcome Gantt Chart !</h3>
          <DnDCalendar
            localizer={localizer}
            defaultDate={moment().toDate()}
            startAccessor="start"
            endAccessor="end"
            events={tasks}
            defaultView="month"
            views={["day", "week", "month", "agenda", "work_week"]}
            style={{ height: 500 }}
            // onSelectEvent={this.handleTaskClick}
            onDoubleClickEvent={this.handleTaskClick}
            onEventDrop={this.moveTask}
            onEventResize={this.resizeTask}
            popup
            resizable
          />
        </div>
        {selectedTask && (
          <div className={styles["task-details-popup"]}>
            <h2>Task Details</h2>
            <p>
              <b>ID: </b>
              {selectedTask.id}
            </p>
            <p>
              <b>Title: </b>
              {selectedTask.title}
            </p>
            <p>
              <b>Start: </b>
              {selectedTask.start.toString()}
            </p>
            <p>
              <b>End: </b>
              {selectedTask.end.toString()}
            </p>
            <button onClick={() => this.setState({ selectedTask: null })}>
              Close
            </button>
            <button onClick={() => this.setState({ selectedTask: null })}>
              Update
            </button>
            <button onClick={() => this.setState({ selectedTask: null })}>
              Delete
            </button>
          </div>
        )}
      </section>
    );
  }
}
