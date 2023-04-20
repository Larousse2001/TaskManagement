import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps } from './IHelloWorldProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class HelloWorld extends React.Component<IHelloWorldProps, {}> {
  public render(): React.ReactElement<IHelloWorldProps> {
    const {
      isDarkTheme,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.helloWorld} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
        </div>
        <div>
          <h3>Welcome to Task Management!</h3>
          <p>
            A web-based task management application that allows users to create, edit, and track tasks, as well as prioritize and categorize them by status.
            The application should feature intuitive drag and drop functionality, and a streamlined user interface that enables users to view tasks in multiple ways,
            including a calendar view, a Gantt chart, and a list view. The application should integrate with existing project management tools and be accessible via
            SharePoint & Teams.
          </p>
          <h4>Application features:</h4>
          <ul className={styles.links}>
            <li><b>Dashboard/Homepage:</b> The dashboard or homepage is the first screen users see after logging in. It could display a summary of their tasks, upcoming deadlines, and other relevant information. This interface could also have links to different views of their tasks.</li>
            <li><b>Task Creation:</b> The task creation interface would allow users to create new tasks by entering a title, description, due date, priority, and status. Users could also assign the task to themselves or another user.</li>
            <li><b>Task List View:</b> The task list view displays a list of all tasks, sorted by due date or priority. Users could filter or sort the tasks based on their preferences, and they could also perform bulk actions like editing or deleting multiple tasks.</li>
            <li><b>Task Detail View:</b> The task detail view shows all the details of a specific task, including its title, description, due date, priority, and status. Users could edit or delete the task from this interface.</li>
            <li><b>Calendar View:</b> The calendar view displays all tasks in a monthly or weekly calendar format. Users could easily view tasks that are due on a specific day or week, and they could also drag and drop tasks to different dates.</li>
            <li><b>Gantt Chart View:</b> The Gantt chart view shows a visual representation of all tasks and their timelines. This interface allows users to easily see how tasks are progressing and identify potential scheduling conflicts.</li>
            <li><b>User Management:</b> The user management interface allows administrators to add, remove, or edit user accounts. It could also enable administrators to assign different roles and permissions to users.</li>
          </ul>
        </div>
        {/* <div id="spListContainer" /> */}
      </section>
    );
  }
}
