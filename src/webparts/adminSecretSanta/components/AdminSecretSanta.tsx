import * as React from 'react';
import styles from './AdminSecretSanta.module.scss';
import { IAdminSecretSantaProps } from './IAdminSecretSantaProps';
import DataService from '../../../Services/DataService';
import {
  DetailsList,
  Selection,
  IColumn,
  buildColumns,
  IColumnReorderOptions,
  IDragDropEvents,
  IDragDropContext
} from 'office-ui-fabric-react/lib/DetailsList';
import { DefaultButton } from 'office-ui-fabric-react';
import { getTheme, mergeStyles } from 'office-ui-fabric-react/lib/Styling';
import { CSVLink, CSVDownload } from "react-csv";


const theme = getTheme();
const dragEnterClass = mergeStyles({
  backgroundColor: theme.palette.neutralLight
});

export interface IAdminSecretSantaState {
  items: any;
  csvData: any;
}

export default class AdminSecretSanta extends React.Component<IAdminSecretSantaProps, IAdminSecretSantaState> {
  public DataService: DataService;
  private _selection: Selection;
  private _dragDropEvents: IDragDropEvents;
  private _draggedItem: undefined;
  private _draggedIndex: number;
  constructor(props: IAdminSecretSantaProps, state: IAdminSecretSantaState) {
    super(props);
    this.DataService = new DataService();
    this._selection = new Selection();
    this._dragDropEvents = this._getDragDropEvents();
    this._draggedIndex = -1;
    this.DataService.getEmployeeList(this.props.lists).then((val) => {
      if (val === true) {
        this.setState({
          items: this.DataService.employeeDetails
        });
      }
    }).catch((error) => {
      console.log(error);
    });
    this.state = {
      items: this.DataService.employeeDetails ? this.DataService.employeeDetails : [],
      csvData: []
    };
  }

  private _getDragDropEvents(): IDragDropEvents {
    return {
      canDrop: (dropContext?: IDragDropContext, dragContext?: IDragDropContext) => {
        return true;
      },
      canDrag: (item?: any) => {
        return true;
      },
      onDragEnter: (item?: any, event?: DragEvent) => {
        // return string is the css classes that will be added to the entering element.
        return dragEnterClass;
      },
      onDragLeave: (item?: any, event?: DragEvent) => {
        return;
      },
      onDrop: (item?: any, event?: DragEvent) => {
        if (this._draggedItem) {
          this._insertBeforeItem(item);
        }
      },
      onDragStart: (item?: any, itemIndex?: number, selectedItems?: any[], event?: MouseEvent) => {
        this._draggedItem = item;
        this._draggedIndex = itemIndex!;
      },
      onDragEnd: (item?: any, event?: DragEvent) => {
        this._draggedItem = undefined;
        this._draggedIndex = -1;
      }
    };
  }

  private _insertBeforeItem(item: any): void {
    const draggedItems = this._selection.isIndexSelected(this._draggedIndex)
      ? (this._selection.getSelection() as any)
      : [this._draggedItem!];

    const items = this.state.items.filter(itm => draggedItems.indexOf(itm) === -1);
    let insertIndex = items.indexOf(item);

    // if dragging/dropping on itself, index will be 0.
    if (insertIndex === -1) {
      insertIndex = 0;
    }

    items.splice(insertIndex, 0, ...draggedItems);

    this.setState({ items: items });
  }

  public columnsAllocatedSanta: IColumn[] = [{
    key: 'asssigned',
    name: 'Allocated Santa Name',
    minWidth: 40,
    fieldName: 'asssigned'
  }];

  public columnsNames: IColumn[] = [{
    key: 'name',
    name: 'Name',
    minWidth: 40,
    fieldName: 'name'
  }];

  public allowDrop(ev) {
    ev.preventDefault();
  }

  public drag(ev) {
    ev.dataTransfer.setData("text", ev.target.id);
  }

  public drop(ev) {
    ev.preventDefault();
    var data = ev.dataTransfer.getData("text");
    ev.target.appendChild(document.getElementById(data));
  }

  public randomClick(): void {
    let tempValue = this.state.items;
    this.DataService.setIsSantaFalse(tempValue);
    let tempDetails: any = this.DataService.generateSecrateSantaAssignedArray(tempValue);
    this.setState({
      items: tempDetails
    });
  }

  public render(): React.ReactElement<IAdminSecretSantaProps> {
    return (
      <div className={styles.adminSecretSanta}>
        <div className={styles.container}>
          {this.props.adminUserEmail.toLowerCase() === this.props.context.pageContext.user.email.toLowerCase() && <div className={styles.row}>
            <h2 style={{ textAlign: "center" }}>welcome Admin of Secret Santa Event</h2>
            <h3 style={{ textAlign: "center" }}>Follow the below steps to allocate Secret Santa</h3>
            <ol>
              <li><div>Click on Randomize allocation - This would populate the Santa against all the users</div>
                <div>we can re-run this process untill we are satisfied with the santa allocation</div>
                <li>Once the allocation of santa is finalized click on Send Email this would send email to all the users</li>
                <li>Click on Download SecretSanta Mapping Sheet for exporting the mapping file in CSV and using it as reference</li>
              </li>
            </ol>
            <div className={styles.row}>
              <div className={styles.column}>
                <DefaultButton text="Randamize allocation" onClick={() => this.randomClick()} allowDisabledFocus />
                <DefaultButton text="Send Email" onClick={() => this.DataService.sendEmail()} allowDisabledFocus style={{ marginLeft: 10 }} />
              </div>
              <div className={styles.column}>
                <CSVLink data={this.state.csvData} filename={"SecretSantaMapping.csv"}
                  onClick={() => {
                    let csvData = this.state.csvData;
                    csvData.splice(0, csvData.length);
                    csvData.push(["Name", "Allocated Santa Name"]);
                    this.DataService.employeeDetails.map((val, index) => {
                      let item = [val.name, this.state.items[index].asssigned];
                      csvData.push(item);
                    });
                    this.setState({
                      csvData: csvData
                    });
                  }}
                >Download SecretSanta Mapping Sheet</CSVLink>
              </div>
            </div>
            <div className={styles.row}>
              <div className={styles.column}>
                <DetailsList
                  items={this.DataService.employeeDetails}
                  columns={this.columnsNames}
                ></DetailsList>
              </div>
              <div className={styles.column}>
                <DetailsList
                  items={this.state.items}
                  columns={this.columnsAllocatedSanta}
                  selection={this._selection}
                  dragDropEvents={this._dragDropEvents}></DetailsList>
              </div>
            </div>
          </div>
          }
        </div>
        <div className={styles.container}>
          {this.props.adminUserEmail.toLowerCase() !== this.props.context.pageContext.user.email.toLowerCase() &&
            <div>{`Sorry You do Not Have access`}</div>
          }
        </div>
      </div>
    );
  }
}
