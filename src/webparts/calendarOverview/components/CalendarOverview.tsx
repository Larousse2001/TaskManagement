/* eslint-disable @typescript-eslint/no-empty-function */
/* eslint-disable no-void */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable react/self-closing-comp */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @rushstack/no-new-null */
import * as React from "react";
import styles from "./CalendarOverview.module.scss";
import { ICalendarOverviewProps } from "./ICalendarOverviewProps";
import { escape } from "@microsoft/sp-lodash-subset";

import { Calendar, momentLocalizer } from "react-big-calendar";
import "react-big-calendar/lib/css/react-big-calendar.css";
import moment from "moment";

import withDragAndDrop from "react-big-calendar/lib/addons/dragAndDrop";
// import 'react-big-calendar/lib/addons/dragAndDrop/styles.scss';

import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import Multiselect from "multiselect-react-dropdown";

export interface Event {
  id: number;
  title: string;
  start: Date;
  end: Date;
  priority: string;
  status: string;
  assignedto: string;
}

const localizer = momentLocalizer(moment);
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const DnDCalendar = withDragAndDrop(Calendar as any);

let x: string = "x"; // priority variable
let y: string = "y"; // status variable

export default class CalendarOverview extends React.Component<
  ICalendarOverviewProps,
  {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    selectedTask: any | null;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any, @rushstack/no-new-null
    selectedButton: any | null;
    showLegend: any | null;
    events: Event[];
    eventsLoaded: true | false;
    options: any[];
    selectedValue: any[];
  }
> {
  constructor(props: ICalendarOverviewProps) {
    super(props);
    this.state = {
      showLegend: null,
      selectedTask: null,
      selectedButton: null, // State to manage popup visibility
      events: [] as Event[],
      eventsLoaded: false,
      options: [],
      selectedValue: [],
    };
  }

  public async componentDidMount(): Promise<void> {
    // eslint-disable-next-line no-void
    void this.getTasks();
    await this.loadEvents();
    const options = await this.selectOptions();
    this.setState({ options });
  }

  handleTaskClick = (event: any) => {
    this.setState({
      selectedTask: event,
    });
  };

  moveTask = (event: any) => {
    this.setState((prevState) => ({
      selectedTask: {
        ...prevState.selectedTask,
        start: event.start,
        end: event.end,
      },
    }));
  };

  handleButtonClick = () => {
    this.setState({
      selectedButton: true, // Open popup when button is clicked
    });
  };

  loadEvents = async () => {
    try {
      const currentTeam = await this.getTeamForCurrentUser();
      const items = await sp.web.lists
        .getByTitle("Tasks")
        .items.filter(`Team eq '${currentTeam}'`)
        .get();

      const events = items.map((item) => ({
        id: item.ID,
        title: item.Title,
        start: new Date(item.StartDate0),
        end: new Date(item.EndDate0),
        priority: item.Priority,
        status: item.Status,
        assignedto: item.AssignedTo0,
      }));

      this.setState({ events, eventsLoaded: true });
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

  // Get filtered items by status
  private getFilteredItemsByStatus = async (status: string): Promise<any[]> => {
    try {
      if (status === "all") {
        void this.getTasks();
      } else if (x !== "x" || x !== "x") {
        if (x === "importanturg") {
          const currentTeam = await this.getTeamForCurrentUser();
          const items: any[] = await sp.web.lists
            .getByTitle("Tasks")
            .items.filter(
              `Team eq '${currentTeam}' and (Priority eq 'Urgent' or Priority eq 'Important' or Priority eq 'Important and urgent') and Status eq '${status}'`
            )
            .get();
          const events: any[] = [];

          items.forEach((item: any) => {
            const task: any = {
              id: item.ID,
              title: item.Title,
              start: new Date(item.StartDate0),
              end: new Date(item.EndDate0),
              priority: item.Priority,
              status: item.Status,
              assignedto: item.AssignedTo0,
            };
            events.push(task);
          });

          this.setState({ events, eventsLoaded: true });
        } else {
          const currentTeam = await this.getTeamForCurrentUser();
          const items: any[] = await sp.web.lists
            .getByTitle("Tasks")
            .items.filter(
              `Team eq '${currentTeam}' and Status eq '${status}' and Priority eq '${x}'`
            )
            .get();
          const events: any[] = [];

          items.forEach((item: any) => {
            const task: any = {
              id: item.ID,
              title: item.Title,
              start: new Date(item.StartDate0),
              end: new Date(item.EndDate0),
              priority: item.Priority,
              status: item.Status,
              assignedto: item.AssignedTo0,
            };
            events.push(task);
          });

          this.setState({ events, eventsLoaded: true });
        }
      } else {
        const currentTeam = await this.getTeamForCurrentUser();
        const items = await sp.web.lists
          .getByTitle("Tasks")
          .items.filter(`Team eq '${currentTeam}' and Status eq '${status}'`)
          .get();

        const events: any[] = [];

        items.forEach((item: any) => {
          const task: any = {
            id: item.ID,
            title: item.Title,
            start: new Date(item.StartDate0),
            end: new Date(item.EndDate0),
            priority: item.Priority,
            status: item.Status,
            assignedto: item.AssignedTo0,
          };
          events.push(task);
        });

        this.setState({ events, eventsLoaded: true });
      }
    } catch (e) {
      console.error(e);
      // Return an empty array in case of any error
      return [];
    }
  };

  // Get filtered items by priority
  private getFilteredItemsByPriority = async (
    priority: string
  ): Promise<any[]> => {
    try {
      if (priority === "all") {
        void this.getTasks();
      } else if (priority === "importanturg") {
        const currentTeam = await this.getTeamForCurrentUser();
        const items: any[] = await sp.web.lists
          .getByTitle("Tasks")
          .items.filter(
            `Team eq '${currentTeam}' and (Priority eq 'Urgent' or Priority eq 'Important' or Priority eq 'Important and urgent')`
          )
          .get();
        const events: any[] = [];

        items.forEach((item: any) => {
          const task: any = {
            id: item.ID,
            title: item.Title,
            start: new Date(item.StartDate0),
            end: new Date(item.EndDate0),
            priority: item.Priority,
            status: item.Status,
            assignedto: item.AssignedTo0,
          };
          events.push(task);
        });

        this.setState({ events, eventsLoaded: true });
      } else if (y !== "y") {
        if (y !== "all") {
          const currentTeam = await this.getTeamForCurrentUser();
          const items: any[] = await sp.web.lists
            .getByTitle("Tasks")
            .items.filter(
              `Team eq '${currentTeam}' and Priority eq '${priority}'`
            )
            .get();
          const events: any[] = [];

          items.forEach((item: any) => {
            const task: any = {
              id: item.ID,
              title: item.Title,
              start: new Date(item.StartDate0),
              end: new Date(item.EndDate0),
              priority: item.Priority,
              status: item.Status,
              assignedto: item.AssignedTo0,
            };
            events.push(task);
          });

          this.setState({ events, eventsLoaded: true });
        } else {
          const currentTeam = await this.getTeamForCurrentUser();
          const items: any[] = await sp.web.lists
            .getByTitle("Tasks")
            .items.filter(
              `Team eq '${currentTeam}' and Priority eq '${priority}' and Status eq '${y}'`
            )
            .get();
          const events: any[] = [];

          items.forEach((item: any) => {
            const task: any = {
              id: item.ID,
              title: item.Title,
              start: new Date(item.StartDate0),
              end: new Date(item.EndDate0),
              priority: item.Priority,
              status: item.Status,
              assignedto: item.AssignedTo0,
            };
            events.push(task);
          });

          this.setState({ events, eventsLoaded: true });
        }
      } else {
        const currentTeam = await this.getTeamForCurrentUser();
        const items: any[] = await sp.web.lists
          .getByTitle("Tasks")
          .items.filter(
            `Team eq '${currentTeam}' and Priority eq '${priority}'`
          )
          .get();
        const events: any[] = [];

        items.forEach((item: any) => {
          const task: any = {
            id: item.ID,
            title: item.Title,
            start: new Date(item.StartDate0),
            end: new Date(item.EndDate0),
            priority: item.Priority,
            status: item.Status,
            assignedto: item.AssignedTo0,
          };
          events.push(task);
        });

        this.setState({ events, eventsLoaded: true });
      }
    } catch (e) {
      console.error(e);
      // Return an empty array in case of any error
      return [];
    }
  };

  // Get Tasks from SharePoint
  private getTasks = async (): Promise<any[]> => {
    try {
      const currentTeam = await this.getTeamForCurrentUser();
      const items: any[] = await sp.web.lists
        .getByTitle("Tasks")
        .items.filter(`Team eq '${currentTeam}'`)
        .get();

      const events: any[] = [];

      items.forEach((item: any) => {
        const task: any = {
          id: item.ID,
          title: item.Title,
          start: new Date(item.StartDate0),
          end: new Date(item.EndDate0),
          priority: item.Priority,
          status: item.Status,
          assignedto: item.AssignedTo0,
        };

        let isDuplicate = false;
        for (let i = 0; i < events.length; i++) {
          if (JSON.stringify(events[i]) === JSON.stringify(task)) {
            isDuplicate = true;
            break;
          }
        }

        if (!isDuplicate) {
          events.push(task);
        }
      });

      this.setState({ events, eventsLoaded: true });
    } catch (e) {
      console.error(e);
      return [];
    }
  };

  private formatDate: any = (date: string): string => {
    const isoDate = new Date(date).toISOString().substring(0, 10);
    const dateParts = isoDate.split("-");
    return `${dateParts[0]}-${dateParts[1]}-${dateParts[2]}`;
  };

  // Select Options List
  private selectOptions = async (): Promise<any[]> => {
    try {
      // Call the getAllItems() function to retrieve the list of employees' full names
      const employeesFullNames = await this.getListEmployees();
      const options: object[] = [];
      for (let index = 0; index < employeesFullNames.length; index++) {
        options.push({ key: employeesFullNames[index], id: index });
      }
      return options;
    } catch (error) {
      console.error(error);
    }
  };

  private handleStatusChange = async (
    event: React.ChangeEvent<HTMLSelectElement>
  ) => {
    const status = event.target.value;
    y = event.target.value;
    // eslint-disable-next-line no-void
    void this.getFilteredItemsByStatus(status);
  };

  private handlePriorityChange = async (
    event: React.ChangeEvent<HTMLSelectElement>
  ) => {
    const priority = event.target.value;
    console.log(x);
    x = event.target.value;
    console.log(x);
    // eslint-disable-next-line no-void
    void this.getFilteredItemsByPriority(priority);
  };

  private readItem = async () => {
    this.handleButtonClick();
    // eslint-disable-next-line no-void
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    this.getItem();
  };

  render() {
    const { hasTeamsContext } = this.props;

    const { selectedTask } = this.state;
    const { selectedButton } = this.state;

    const { eventsLoaded, events } = this.state;

    const { showLegend } = this.state;

    const handleLegendClick = (event: any) => {
      if (showLegend === true) {
        this.setState({
          showLegend: null,
        });
      } else {
        this.setState({
          showLegend: true,
        });
      }
    };

    return (
      <section
        className={`${styles.calendarOverview} ${
          hasTeamsContext ? styles.teams : ""
        }`}
      >
        <div className={styles.welcome}>
          <h2>TaskMaster: Take Control of Your Schedule!</h2>
        </div>
        <br></br>
        <div className="special">
          <div className={styles.buttonSec}>
            <div className={styles.itemF}>
              <select
                className={styles.inputField}
                id="status"
                onChange={this.handleStatusChange}
              >
                <option value="all">All</option>
                <option value="Active">Active</option>
                <option value="Completed">Completed</option>
                <option value="On Hold">On Hold</option>
                <option value="Cancelled">Cancelled</option>
              </select>
              <div className={styles.fieldL}>
                <b> Filter By Status</b>
              </div>
            </div>
            <div className={styles.itemF}>
              <select
                className={styles.inputField}
                id="priority"
                onChange={this.handlePriorityChange}
              >
                <option value="all">All</option>
                <option value="Important">Important</option>
                <option value="importanturg">Important and urgent</option>
                <option value="Urgent">Urgent</option>
                <option value="Neither">Neither</option>
              </select>
              <div className={styles.fieldL}>
                <b> Filter By Priority</b>
              </div>
            </div>
            <div className={styles.button}>
              <span className={styles.label} onClick={handleLegendClick}>
                Legend
              </span>
            </div>
          </div>
        </div>
        {showLegend && (
          <div className={styles.legend}>
            <div>
              <span style={{ color: "RGBA(0, 169, 235, 1)" }}>
                <b>Neither</b>
              </span>
            </div>
            <div>
              <span style={{ color: "RGBA(254, 211, 76, 1)" }}>
                <b>Important</b>
              </span>
            </div>
            <div>
              <span style={{ color: "RGBA(255, 153, 18, 1)" }}>
                <b>Urgent</b>
              </span>
            </div>
            <div>
              <span style={{ color: "RGBA(250, 0, 87, 1)" }}>
                <b>Important and urgent</b>
              </span>
            </div>
          </div>
        )}
        <br></br>
        <br></br>
        <div>
          {!eventsLoaded ? (
            <div>
              <b>Loading...</b>
            </div>
          ) : (
            <DnDCalendar
              localizer={localizer}
              defaultDate={moment().toDate()}
              startAccessor="start"
              endAccessor="end"
              events={events}
              defaultView="month"
              views={["day", "week", "month", "agenda"]}
              style={{ height: 500 }}
              onDoubleClickEvent={this.handleTaskClick}
              onEventDrop={this.moveTask}
              popup
              eventPropGetter={(event: Event) => {
                const style: React.CSSProperties = {
                  backgroundColor: "RGBA(0, 169, 235, 1)",
                };

                if (event.priority === "Important") {
                  style.backgroundColor = "RGBA(254, 211, 76, 1)";
                }

                if (event.priority === "Urgent") {
                  style.backgroundColor = "RGBA(255, 153, 18, 1)";
                }

                if (event.priority === "Important and urgent") {
                  style.backgroundColor = "RGBA(250, 0, 87, 1)";
                }

                return {
                  style,
                };
              }}
            />
          )}
        </div>
        <br></br>
        <div className={styles.buttonSection}>
          <div className={styles.button}>
            <span className={styles.label} onClick={this.handleButtonClick}>
              Add a Task 📝
            </span>
          </div>
        </div>
        {selectedButton && (
          <div className={styles.teamOv}>
            <form onSubmit={(event) => event.preventDefault()}>
              <div className={styles.itemField}>
                <div className={styles.fieldLabel}>Title</div>
                <input
                  className={styles.fieldInput}
                  type="text"
                  id="title"
                ></input>
              </div>
              <div className={styles.itemField}>
                <div className={styles.fieldLabel}>Assigned To</div>
                <Multiselect
                  id="assignedto"
                  options={this.state.options} // Options to display in the dropdown
                  selectedValues={this.state.selectedValue} // Preselected value to persist in dropdown
                  onKeyPressFn={function noRefCheck() {}}
                  onRemove={function noRefCheck() {}}
                  onSearch={function noRefCheck() {}}
                  onSelect={(selectedList, selectedItem) => {
                    this.setState({ selectedValue: selectedList });
                  }}
                  displayValue="key" // Property name to display in the dropdown options
                  showCheckbox
                  placeholder="Choose the member"
                  // className={styles.multiselect}
                />
              </div>
              <div className={styles.itemField}>
                <div className={styles.fieldLabel}>Start Date</div>
                <input
                  className={styles.dateInput}
                  type="date"
                  id="startdate"
                />
              </div>
              <div className={styles.itemField}>
                <div className={styles.fieldLabel}>End Date</div>
                <input className={styles.dateInput} type="date" id="enddate" />
              </div>
              <div className={styles.itemField}>
                <div className={styles.fieldLabel}>Status</div>
                <select className={styles.selectField} id="stat">
                  <option value="Active">Active</option>
                  <option value="Completed">Completed</option>
                  <option value="On Hold">On Hold</option>
                  <option value="Cancelled">Cancelled</option>
                </select>
              </div>
              <div className={styles.itemField}>
                <div className={styles.fieldLabel}>Priority</div>
                <select className={styles.selectField} id="pri">
                  <option value="Important">Important</option>
                  <option value="Important and urgent">
                    Important and urgent
                  </option>
                  <option value="Urgent">Urgent</option>
                  <option value="Neither">Neither</option>
                </select>
              </div>
              <div className={styles.buttonSection}>
                <div className={styles.button}>
                  <span className={styles.label} onClick={this.createItem}>
                    Create
                  </span>
                </div>
                <div className={styles.button}>
                  <span className={styles.label} onClick={this.updateItem}>
                    Update
                  </span>
                </div>
                <div className={styles.button}>
                  <span
                    className={styles.label}
                    onClick={() => this.setState({ selectedButton: null })}
                  >
                    Cancel
                  </span>
                </div>
              </div>
            </form>
          </div>
        )}
        {selectedTask && (
          <div className={styles["task-details-popup"]}>
            <h2>Task Details</h2>
            <div className={styles.itemField}>
              <div className={styles.fieldLabel}>ID:</div>
              <input
                type="text"
                id="itemID"
                value={selectedTask.id}
                disabled
              ></input>
            </div>
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
            <p>
              <b>Assigned To: </b>
              {selectedTask.assignedto}
            </p>
            <p>
              <b>Status: </b>
              {selectedTask.status}
            </p>
            <p>
              <b>Priority: </b>
              {selectedTask.priority}
            </p>
            <button onClick={this.readItem}>Update</button>
            <button onClick={this.deleteItem}>Delete</button>
            <button onClick={() => this.setState({ selectedTask: null })}>
              Close
            </button>
          </div>
        )}
      </section>
    );
  }

  //Create Item
  private createItem = async () => {
    try {
      const currentTeam = await this.getTeamForCurrentUser();

      const startDate = (
        document.getElementById("startdate") as HTMLInputElement
      ).value;
      const endDate = (document.getElementById("enddate") as HTMLInputElement)
        .value;

      if (endDate < startDate) {
        alert("End date must be greater than start date.");
        return;
      }

      const addItem = await sp.web.lists.getByTitle("Tasks").items.add({
        Title: (document.getElementById("title") as HTMLInputElement).value,
        AssignedTo0: this.state.selectedValue
          .map((item) => item.key)
          .join(", "),
        StartDate0: (document.getElementById("startdate") as HTMLInputElement)
          .value,
        EndDate0: (document.getElementById("enddate") as HTMLInputElement)
          .value,
        Status: (document.getElementById("stat") as HTMLInputElement).value,
        Priority: (document.getElementById("pri") as HTMLInputElement).value,
        Team: currentTeam,
      });
      alert(`Item created successfully with ID: ${addItem.data.ID}`);
      void this.getTasks();
      await this.loadEvents();
      this.setState({ selectedButton: null });
    } catch (e) {
      console.error(e);
    }
  };

  //Get Item by ID
  private getItem = async () => {
    try {
      const id: number = (document.getElementById("itemID") as HTMLInputElement)
        .value as unknown as number;
      console.log(id);
      if (id > 0) {
        const item: any = await sp.web.lists
          .getByTitle("Tasks")
          .items.getById(id)
          .get();
        (document.getElementById("title") as HTMLInputElement).value =
          item.Title;
        (document.getElementById("satus") as HTMLInputElement).value =
          item.Status;
        (document.getElementById("priority") as HTMLInputElement).value =
          item.Priority;
        (document.getElementById("startdate") as HTMLInputElement).value =
          this.formatDate(item.StartDate0);
        (document.getElementById("enddate") as HTMLInputElement).value =
          this.formatDate(item.EndDate0);
        (document.getElementById("assignedto") as HTMLInputElement).value =
          item.AssignedTo0;
      } else {
        alert(`Please enter a valid item id.`);
      }
    } catch (e) {
      console.error(e);
    }
  };

  //Update Item
  private updateItem = async () => {
    try {
      const currentTeam = await this.getTeamForCurrentUser();

      const startDate = (
        document.getElementById("startdate") as HTMLInputElement
      ).value;
      const endDate = (document.getElementById("enddate") as HTMLInputElement)
        .value;

      if (endDate < startDate) {
        alert("End date must be greater than start date.");
        return;
      }

      const id: number = (document.getElementById("itemID") as HTMLInputElement)
        .value as unknown as number;
      if (id > 0) {
        const itemUpdate = await sp.web.lists
          .getByTitle("Tasks")
          .items.getById(id)
          .update({
            Title: (document.getElementById("title") as HTMLInputElement).value,
            AssignedTo0: this.state.selectedValue
              .map((item) => item.key)
              .join(", "),
            StartDate0: (
              document.getElementById("startdate") as HTMLInputElement
            ).value,
            EndDate0: (document.getElementById("enddate") as HTMLInputElement)
              .value,
            Status: (document.getElementById("status") as HTMLInputElement)
              .value,
            Priority: (document.getElementById("priority") as HTMLInputElement)
              .value,
            Team: currentTeam,
          });
        alert(`Item with ID: ${id} updated successfully!`);
        void this.getTasks();
        await this.loadEvents();
        this.setState({ selectedTask: null });
      } else {
        alert(`Please enter a valid item id.`);
      }
    } catch (e) {
      console.error(e);
    }
  };

  //Delete Item
  private deleteItem = async () => {
    try {
      const id: number = parseInt(
        (document.getElementById("itemID") as HTMLInputElement).value
      );
      if (id > 0) {
        const deleteItem = await sp.web.lists
          .getByTitle("Tasks")
          .items.getById(id)
          .delete();
        console.log(deleteItem);
        // eslint-disable-next-line @typescript-eslint/no-floating-promises
        this.getTasks();
        alert(`Item ID: ${id} deleted successfully!`);
      } else {
        alert(`Please enter a valid item id.`);
      }
    } catch (e) {
      console.error(e);
    }
  };
}
