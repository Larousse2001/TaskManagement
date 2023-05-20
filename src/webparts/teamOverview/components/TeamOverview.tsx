/* eslint-disable react/no-unescaped-entities */
/* eslint-disable no-void */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable react/self-closing-comp */
import * as React from "react";
import styles from "./TeamOverview.module.scss";
import { ITeamOverviewProps } from "./ITeamOverviewProps";

import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export default class TeamOverview extends React.Component<
  ITeamOverviewProps,
  {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any, @rushstack/no-new-null
    selectedButton: any | null;
  }
> {
  constructor(props: ITeamOverviewProps) {
    super(props);
    this.state = {
      selectedButton: null, // State to manage popup visibility
    };
  }

  handleButtonClick = () => {
    this.setState({
      selectedButton: true, // Open popup when button is clicked
    });
  };

  handleButtonClose = () => {
    this.setState({
      selectedButton: null, // Close popup when close button is clicked
    });
  };

  componentDidMount() {
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    this.getAllItems(); // Call the function when the component is mounted
  }

  // Get all items
  getAllItems = async () => {
    try {
      const currentTeam = await this.getTeamForCurrentUser();
      const items = await sp.web.lists
        .getByTitle("Employees")
        .items.filter(`Team eq '${currentTeam}'`)
        .get();
      if (items.length > 0) {
        let html = `
        <table style="width:100%;border-collapse:collapse;">
          <tr style="border-bottom:1px solid #000;">
            <th style="text-align:left;padding:16px;">Full Name</th>
            <th style="text-align:left;padding:16px;">Email</th>
          </tr>
      `;
        items.map((item, index) => {
          html += `
          <tr style="border-bottom:1px solid #000;cursor:pointer;">
            <td style="text-align:left;padding:16px;">${item.Fullname}</td>
            <td style="text-align:left;padding:16px;">${item.Email}</td>
          </tr>
        `;
          setTimeout(() => {
            const tr = document.querySelector(
              `#allItems tr:nth-child(${index + 2})`
            );
            if (tr) {
              tr.addEventListener("click", this.displayPopup.bind(this, item));
            }
          }, 100);
        });
        html += `</table>`;
        document.getElementById("allItems").innerHTML = html;
      } else {
        alert(`List is empty.`);
      }
    } catch (e) {
      console.error(e);
    }
  };

  private displayPopup = (item: any) => {
    const popupHtml = `
    <div style="position: absolute;top: 95px;left: 50%;transform: translateX(-50%);background-color: #ffffff;border-radius: 5px;
      box-shadow: 0 0 10px rgba(0, 0, 0, 0.2);padding: 20px;max-width: 400px;width: 100%;max-height: 80vh;overflow-y: auto;z-index: 9999;">
      <h2 style="font-size: 24px;margin-top: 0;">Employee Details</h2>
      <p style="margin-bottom: 10px;">
        <b>ID: </b>
        <input
          type="text"
          id="itemID"
          value=${item.ID}
          disabled
        ></input>
      </p>
      <p style="margin-bottom: 10px;">
        <b>Fullname: </b>
        ${item.Fullname}
      </p>
      <p style="margin-bottom: 10px;">
        <b>Email: </b>
        ${item.Email}
      </p>
      <p style="margin-bottom: 10px;">
        <b>Role: </b>
        ${item.Role}
      </p>
      <button  id="deleteButton" style="background-color: #0078d4;color: #ffffff;border: none;padding: 10px 15px;cursor: pointer;font-size: 16px;
        border-radius: 5px;margin-right: 10px;">Delete</button>
      <button id="hide" style="background-color: #0078d4;color: #ffffff;border: none;padding: 10px 15px;cursor: pointer;font-size: 16px;
        border-radius: 5px;margin-right: 10px;">Close</button>
    </div>
    `;
    setTimeout(() => {
      // Add event listener to the delete button
      const deleteButton = document.getElementById("deleteButton");
      if (deleteButton) {
        deleteButton.addEventListener("click", () => {
          void this.deleteItem(item.ID);
        });
      }
      // Add event listener to the delete button
      const hideButton = document.getElementById("hide");
      if (hideButton) {
        hideButton.addEventListener("click", () => {
          void this.hidePopup();
        });
      }
    }, 100);
    document.getElementById("task-details-popup").innerHTML = popupHtml;
  };

  //Delete Item
  private deleteItem = async (itemId: number) => {
    try {
      const deleteItem = await sp.web.lists
        .getByTitle("Employees")
        .items.getById(itemId)
        .delete();
      console.log(deleteItem);
      alert(`Item ID: ${itemId} deleted successfully!`);
      this.hidePopup();
      void this.getAllItems();
    } catch (e) {
      console.error(e);
    }
  };

  //Hide the popup
  private hidePopup = () => {
    const popupHtml = `<div></div>`
    document.getElementById("task-details-popup").innerHTML = popupHtml
  };

  // Get all items and return a list filled with employees' emails or an empty list
  private getListEmployees = async (): Promise<string[]> => {
    try {
      const items: any[] = await sp.web.lists
        .getByTitle("Employees")
        .items.get();

      // Create an array to store employees' full names
      const employeesEmails: string[] = [];

      if (items.length > 0) {
        // Loop through the retrieved items and extract the "Fullname" field value
        items.forEach((item) => {
          // Add the "Fullname" field value to the employees' full names array
          employeesEmails.push(item.Email);
        });
      }
      console.log(employeesEmails);
      // Return the employees' full names array
      return employeesEmails;
    } catch (e) {
      console.error(e);
      // Return an empty array in case of any error
      return [];
    }
  };

  public render(): React.ReactElement<ITeamOverviewProps> {
    const { selectedButton } = this.state;

    return (
      <div className={styles.teamOverview}>
        <section>
          <h1 className={styles.title}>TeamView: Teamwork divides the task and multiplies the success!</h1>
          <div id="allItems">
            {/* Render the content generated by getAllItems function here */}
          </div>
          <div id="task-details-popup">
            {/* Render the content generated by displayPopup function here */}
          </div>
          <h1 className={styles.title}>....</h1>
        </section>
        <section>
          <div className={styles.buttonSection}>
            <div className={styles.button}>
              <span className={styles.label} onClick={this.handleButtonClick}>
                Add Member 🤵
              </span>
            </div>
          </div>
          {selectedButton && (
            <div className={styles.teamOv}>
              <form onSubmit={(event) => event.preventDefault()}>
                <div className={styles.itemField}>
                  <div className={styles.fieldLabel}>Full Name</div>
                  <input className={styles.fieldInput} type="text" id="fullName"></input>
                </div>
                <div className={styles.itemField}>
                  <div className={styles.fieldLabel}>Email</div>
                  <input className={styles.fieldInput} type="text" id="email"></input>
                </div>
                <div className={styles.itemField}>
                  <div className={styles.fieldLabel}>Role</div>
                  <select className={styles.selectField} id="role">
                    <option value="Agent">Agent</option>
                    <option value="Team Lead">Team Lead</option>
                  </select>
                </div>
                <div className={styles.buttonSection}>
                  <div className={styles.button}>
                    <span className={styles.label} onClick={this.createItem}>
                      Create
                    </span>
                  </div>
                  <div className={styles.button}>
                    <span
                      className={styles.label}
                      onClick={this.handleButtonClose}
                    >
                      Cancel
                    </span>
                  </div>
                </div>
              </form>
            </div>
          )}
        </section>
      </div>
    );
  }

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

  //Create Item
  private createItem = async () => {
    try {
      const currentTeam = await this.getTeamForCurrentUser();
      const email = (document.getElementById("email") as HTMLInputElement).value;
      const employeesEmails = await this.getListEmployees();
      const existingItems = employeesEmails.filter((item) => item === email);
      if (existingItems.length > 0) {
        alert(`Item with email ${email} already exists in another team.`);
        return;
      }
      const addItem = await sp.web.lists.getByTitle("Employees").items.add({
        Fullname: (document.getElementById("fullName") as HTMLInputElement).value,
        Email: email,
        Role: (document.getElementById("role") as HTMLInputElement).value,
        Team: currentTeam,
      });
      void this.getAllItems();
      alert(`Item created successfully with ID: ${addItem.data.ID}`);
    } catch (e) {
      console.error(e);
    }
  };
}
