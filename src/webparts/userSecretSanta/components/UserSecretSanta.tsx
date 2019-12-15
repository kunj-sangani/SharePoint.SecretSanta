import * as React from 'react';
import styles from './UserSecretSanta.module.scss';
import { IUserSecretSantaProps } from './IUserSecretSantaProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import { DefaultButton, Dialog, DialogFooter, PrimaryButton, DialogType, getId } from 'office-ui-fabric-react';
import * as moment from "moment";
import DataService from '../../../Services/DataService';

const santa: any = require('../../../assets/santa-claus-animated-gif-4.gif');
const down: any = require('../../../assets/down.jpg');

export interface IUserSecretSantaState {
  myRichText: string;
  eventDate: any;
  differenceDate: any;
  itemID: any;
  hideDialog: boolean;
  nowLoad: boolean;
  differencehour: any;
  differenceminute: any;
  differencesecond: any;
}

export default class UserSecretSanta extends React.Component<IUserSecretSantaProps, IUserSecretSantaState> {
  public DataService: DataService;
  constructor(props: IUserSecretSantaProps, state: IUserSecretSantaState) {
    super(props);
    this.DataService = new DataService();
    this.state = {
      myRichText: "",
      eventDate: moment(this.props.propsdatetime ? this.props.propsdatetime.displayValue : new Date),
      differenceDate: null,
      itemID: null,
      hideDialog: true,
      nowLoad: false,
      differencehour: null,
      differenceminute: null,
      differencesecond: null
    };
    this.tickerTime();
    this.fetchExistingValues();
  }

  private fetchExistingValues = () => {
    this.DataService.fetchExistingWishValue(this.props.context.pageContext.user.email).then((val) => {
      console.log(val);
      if (val.length > 0) {
        this.setState({
          myRichText: val[0].Wish,
          itemID: val[0].ID,
          nowLoad: true
        });
        this.onTextChange(val[0].Wish);
      }
    }).catch((error) => {
      console.log(error);
    });
  }

  private tickerTime() {
    let tempvalue: any = this.state.eventDate.diff(moment(new Date, "DD/MM/YYYY HH:mm:ss"));
    let date: any = moment.utc(tempvalue).format("DD");
    let hour: any = moment.utc(tempvalue).format("HH");
    let minute: any = moment.utc(tempvalue).format("mm");
    let second: any = moment.utc(tempvalue).format("ss");
    this.setState({
      differenceDate: date,
      differencehour: hour,
      differenceminute: minute,
      differencesecond: second
    });
    setTimeout(() => {
      this.tickerTime();
    }, 1000);
  }

  private onTextChange = (newText: string) => {
    this.setState({
      myRichText: newText
    });

    return newText;
  }

  public saveWishList() {
    if (this.state.itemID) {
      this.DataService.addupdateWishValue(this.state.myRichText, this.state.itemID).then((event) => {
        this.setState({
          itemID: event.data.ID,
          hideDialog: false
        });
      }).catch((error) => {
        console.log(error);
      });
    } else {
      this.DataService.addupdateWishValue(this.state.myRichText).then((event) => {
        this.setState({
          itemID: event.data.ID,
          hideDialog: false
        });
      }).catch((error) => {
        console.log(error);
      });
    }
  }

  private _closeDialog = (): void => {
    this.setState({ hideDialog: true });
  }

  private _labelId: string = getId('dialogLabel');
  private _subTextId: string = getId('subTextLabel');

  public render(): React.ReactElement<IUserSecretSantaProps> {
    return (
      <div className={styles.userSecretSanta}>
        <div className={styles.container}>
          <div className={styles.rowWithOutPadding}>
            <h2 className={styles.title}>Welcome {this.props.context.pageContext.user.displayName} to Secret Santa Event 2019!!!</h2>
            <h3 className={styles.subTitle}>Please enter your wish list below so that Santa can fulfil it</h3>
            <img src={down} style={{ float: "left", position: "relative", left: "40%",height:"170px" }}></img>
          </div>{this.state.nowLoad &&
            <RichText value={this.state.myRichText}
              onChange={(text) => this.onTextChange(text)}
            />
          }
          <div className={styles.rowWithLeftRightPadding} style={{paddingTop:10}}>
            <div className={styles.column8}>
              <DefaultButton text="Save Your Wish" onClick={() => this.saveWishList()} allowDisabledFocus />
            </div>
          </div>
          <div className={styles.rowWithLeftRightPadding}>
            <h3 style={{ textAlign: "center" }}>Count Down for Secret Santa Event</h3>
            <img src={santa} style={{ height: 150, padding: 10 }}></img>
            <div className={styles.divCountDown}>
              <div className={styles.countDownTime}>{this.state.differencesecond}</div>
              <div className={styles.countDownText}>Second</div>
            </div>
            <div className={styles.divCountDown}>
              <div className={styles.countDownTime}>{this.state.differenceminute}</div>
              <div className={styles.countDownText}>Minute</div>
            </div>
            <div className={styles.divCountDown}>
              <div className={styles.countDownTime}>{this.state.differencehour}</div>
              <div className={styles.countDownText}>Hour</div>
            </div>
            <div className={styles.divCountDown}>
              <div className={styles.countDownTime}>{this.state.differenceDate}</div>
              <div className={styles.countDownText}>Day</div>
            </div>
          </div>
        </div>
        <Dialog
          hidden={this.state.hideDialog}
          onDismiss={this._closeDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Wish Save Successfully',
            closeButtonAriaLabel: 'Close',
            subText: 'Yaayy wait for your santa to deliver you your gift'
          }}
          modalProps={{
            titleAriaId: this._labelId,
            subtitleAriaId: this._subTextId,
            isBlocking: false,
            styles: { main: { maxWidth: 450 } }
          }}
        >
          <DialogFooter>
            <PrimaryButton onClick={this._closeDialog} text="Close" />
          </DialogFooter>
        </Dialog>
      </div>
    );
  }
}
