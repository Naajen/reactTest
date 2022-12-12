import * as React from 'react';
import styles from './SpfxCrudPnp.module.scss';
import { ISpfxCrudPnpProps } from './ISpfxCrudPnpProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";



export default class SpfxCrudPnp extends React.Component<ISpfxCrudPnpProps, {}> {
  public render(): React.ReactElement<ISpfxCrudPnpProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.spfxCrudPnp} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        <div>
        <div className={styles.spfxCrudPnp}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <div className={styles.itemField}>
                <div className={styles.fieldLabel}>Item ID:</div>
                <input type="text" id='itemId'></input>
              </div>
              <div className={styles.itemField}>
                <div className={styles.fieldLabel}>Full Name</div>
                <input type="text" id='fullName'></input>
              </div>
              <div className={styles.itemField}>
                <div className={styles.fieldLabel}>Age</div>
                <input type="text" id='age'></input>
              </div>
              <div className={styles.itemField}>
                <div className={styles.fieldLabel}>All Items:</div>
                <div id="allItems"></div>
              </div>
              <div className={styles.buttonSection}>
                <div className={styles.button}>
                  <span className={styles.label} onClick={this.createItem}>Create</span>
                </div>
                <div className={styles.button}>
                  <span className={styles.label} onClick={this.getItemById}>Read</span>
                </div>
                <div className={styles.button}>
                  <span className={styles.label} onClick={this.getAllItems}>Read All</span>
                </div>
                <div className={styles.button}>
                  <span className={styles.label} onClick={this.updateItem}>Update</span>
                </div>
                <div className={styles.button}>
                  <span className={styles.label} onClick={this.deleteItem}>Delete</span>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>

          <ul className={styles.links}>
            <li><a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">SharePoint Framework Overview</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank" rel="noreferrer">Use Microsoft Graph in your solution</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank" rel="noreferrer">Build for Microsoft Teams using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank" rel="noreferrer">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank" rel="noreferrer">Publish SharePoint Framework applications to the marketplace</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank" rel="noreferrer">SharePoint Framework API reference</a></li>
            <li><a href="https://aka.ms/m365pnp" target="_blank" rel="noreferrer">Microsoft 365 Developer Community</a></li>
          </ul>
        </div>
      </section>
    );
  }
  //Create Item
  private createItem = async () => {
    try {
      const addItem = await sp.web.lists.getByTitle("EmployeeDetails").items.add({
        'Title': document.getElementById("fullName")['value'],
        'Age': document.getElementById("age")['value']
      });
      console.log(addItem);
      alert(`Item created successfully with ID: ${addItem.data.ID}`);
    }
    catch (e) {
      console.error(e);
    }
  }
 
  //Get Item by ID
  private getItemById = async () => {
    try {
      const id: number = document.getElementById('itemId')['value'];
      if (id > 0) {
        const item: any = await sp.web.lists.getByTitle("EmployeeDetails").items.getById(id).get();
        document.getElementById('fullName')['value'] = item.Title;
        document.getElementById('age')['value'] = item.Age;
      }
      else {
        alert(`Please enter a valid item id.`);
      }
    }
    catch (e) {
      console.error(e);
    }
  }
 
  //Get all items
  private getAllItems = async () => {
    try {
      const items: any[] = await sp.web.lists.getByTitle("EmployeeDetails").items.get();
      console.log(items);
      if (items.length > 0) {
        var html = `<table><tr><th>ID</th><th>Full Name</th><th>Age</th></tr>`;
        items.map((item, index) => {
          html += `<tr><td>${item.ID}</td><td>${item.Title}</td><td>${item.Age}</td></li>`;
        });
        html += `</table>`;
        document.getElementById("allItems").innerHTML = html;
      } else {
        alert(`List is empty.`);
      }
    }
    catch (e) {
      console.error(e);
    }
  }
 
  //Update Item
  private updateItem = async () => {
    try {
      const id: number = document.getElementById('itemId')['value'];
      if (id > 0) {
        const itemUpdate = await sp.web.lists.getByTitle("EmployeeDetails").items.getById(id).update({
          'Title': document.getElementById("fullName")['value'],
          'Age': document.getElementById("age")['value']
        });
        console.log(itemUpdate);
        alert(`Item with ID: ${id} updated successfully!`);
      }
      else {
        alert(`Please enter a valid item id.`);
      }
    }
    catch (e) {
      console.error(e);
    }
  }
 
  //Delete Item
  private deleteItem = async () => {
    try {
      const id: number = parseInt(document.getElementById('itemId')['value']);
      if (id > 0) {
        let deleteItem = await sp.web.lists.getByTitle("EmployeeDetails").items.getById(id).delete();
        console.log(deleteItem);
        alert(`Item ID: ${id} deleted successfully!`);
      }
      else {
        alert(`Please enter a valid item id.`);
      }
    }
    catch (e) {
      console.error(e);
    }
  }

}
