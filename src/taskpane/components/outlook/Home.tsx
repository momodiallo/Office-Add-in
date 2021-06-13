/* eslint-disable no-undef */
import * as React from "react";
import { DefaultButton, Panel } from "office-ui-fabric-react";
import { getLocaleStrings } from "../../loc/getLocaleStrings";
import { IOutlookServices } from "../../../services/outlook/IOutlookServices";
import OutlookServices from "../../../services/outlook/OutlookServices";

export interface HomeProps {
  token: string;
}

export interface HomeState {
  openSSOToken: boolean;
  body: any;
  item: any;
}

export default class Home extends React.Component<HomeProps, HomeState> {
  private readonly _outlookServices: IOutlookServices;

  constructor(props, context) {
    super(props, context);
    this.state = {
      openSSOToken: false,
      body: "",
      item: {},
    };

    this._outlookServices = new OutlookServices();
  }

  componentDidMount() {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, this._loadCurrentEmailInfo.bind(this));
    this._loadCurrentEmailInfo();
  }

  _loadCurrentEmailInfo = async () => {
    let body = await this._outlookServices.getCurrentEmailBody();
    let item = this._outlookServices.getCurrentEmailItem();

    this.setState({
      body: body,
      item: item,
    });
  };

  render() {
    const strings = getLocaleStrings();
    const token = atob(this.props.token.split(".")[1]);
    const emailItem = this.state.item;
    return (
      <div>
        <div className="ms-homeTitle">{strings.Home}</div>
        <div className="ms-homeContent">
          <h1 className="ms-homeHeading">{strings.Heading}</h1>
          <p className="mailInfoTitle">Email Info</p>
          <ul>
            <li>
              <span>From:</span> <p>{emailItem.from}</p>
            </li>
            <li>
              <span>To:</span> <p>{emailItem.to}</p>
            </li>
            <li>
              <span>CC:</span> <p>{emailItem.cc}</p>
            </li>
            <li>
              <span>Subject:</span> <p>{emailItem.subject}</p>
            </li>
            <li>
              <span>Attachement name:</span> <p>{emailItem.attachments}</p>
            </li>
            <li>
              <span>Body:</span> <p>{this.state.body}</p>
            </li>
          </ul>
        </div>
        <DefaultButton className="openSSOButton" onClick={() => this.setState({ openSSOToken: true })}>
          {strings.OpenPanelButton}
        </DefaultButton>
        <Panel
          headerText={strings.PanelTitle}
          isOpen={this.state.openSSOToken}
          onDismiss={this.dismissPanel}
          closeButtonAriaLabel="Close"
        >
          <p>{strings.PanelContent}</p>
          <div>
            <pre>{JSON.stringify(token, null, 2)}</pre>
          </div>
        </Panel>
      </div>
    );
  }

  dismissPanel = () => {
    this.setState({ openSSOToken: false });
  };
}
