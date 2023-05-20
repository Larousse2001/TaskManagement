/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable react/self-closing-comp */
import * as React from "react";
import styles from "./Register.module.scss";
import { IRegisterProps } from "./IRegisterProps";

import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export default class Register extends React.Component<IRegisterProps, {}> {
  public render(): React.ReactElement<IRegisterProps> {
    return (
      <div className={styles.register}>
        <h1 className={styles.title}>Registration for Team Lead 👔</h1>
        <div className={styles.teamOv}>
          <form onSubmit={(event) => event.preventDefault()}>
            <div className={styles.itemField}>
              <div className={styles.fieldLabel}>Full Name</div>
              <input type="text" id="fullName"></input>
            </div>
            <div className={styles.itemField}>
              <div className={styles.fieldLabel}>Email</div>
              <input type="text" id="email"></input>
            </div>
            <div className={styles.itemField}>
              <div className={styles.fieldLabel}>Team Name:</div>
              <input type="text" id="team"></input>
            </div>
            <div className={styles.buttonSection}>
              <div className={styles.button}>
                <span className={styles.label} onClick={this.createItem}>
                  Register
                </span>
              </div>
              <div className={styles.button}>
                <span className={styles.label}>Cancel</span>
              </div>
            </div>
          </form>
        </div>
      </div>
    );
  }

  //Create Item
  private createItem = async () => {
    try {
      const addItem = await sp.web.lists.getByTitle("Employees").items.add({
        Fullname: (document.getElementById("fullName") as HTMLInputElement)
          .value,
        Email: (document.getElementById("email") as HTMLInputElement).value,
        Team: (document.getElementById("team") as HTMLInputElement).value,
        Role: "Team Lead",
      });
      console.log(addItem);
      alert(`Item created successfully with ID: ${addItem.data.ID}`);
    } catch (e) {
      console.error(e);
    }
  };
}
