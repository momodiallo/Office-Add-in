/* eslint-disable no-undef */
import * as React from "react";
import { Panel, TextField, MessageBar, MessageBarType, DefaultButton } from "office-ui-fabric-react";
import { getLocaleStrings } from "../../loc/getLocaleStrings";
import WordServices from "../../../services/word/WordServices";
import { IWordServices } from "../../../services/word/IWordServices";

export interface HomeProps {
  token: string;
}

export interface HomeState {
  projectRef: string;
  openSSOToken: boolean;
  success: boolean;
}

export default class Home extends React.Component<HomeProps, HomeState> {
  private readonly _wordServices: IWordServices;

  constructor(props, context) {
    super(props, context);
    this.state = {
      projectRef: "",
      openSSOToken: false,
      success: false,
    };

    this._wordServices = new WordServices();
  }

  componentDidMount() {
    this._wordServices.getProjectRef();
  }

  render() {
    var strings = getLocaleStrings();
    var token = atob(this.props.token.split(".")[1]);
    return (
      <div>
        <div className="ms-homeTitle">{strings.Home}</div>
        <div className="ms-homeContent">
          <TextField
            label="Lexor project ref:"
            placeholder="Enter the lexor project ref"
            value={this.state.projectRef}
            onChange={this._onChange}
          />
          <DefaultButton className="updateButton" onClick={this.updateProjectRef}>
            {strings.UpdateButton}
          </DefaultButton>
          <div>
            <h1 className="ms-homeHeading">{strings.Heading}</h1>
            <p className="content">{strings.HomeContent}</p>
          </div>
          <DefaultButton className="openSSOButton" onClick={() => this.setState({ openSSOToken: true })}>
            {strings.OpenPanelButton}
          </DefaultButton>
        </div>
        {this.state.success && (
          <MessageBar
            className="messageBar"
            onDismiss={() => this.setState({ success: false })}
            messageBarType={MessageBarType.success}
            isMultiline={false}
          >
            {strings.MessageBarSuccess}
          </MessageBar>
        )}
        <Panel
          headerText={strings.PanelTitle}
          isOpen={this.state.openSSOToken}
          onDismiss={this.dismissPanel}
          closeButtonAriaLabel="Close"
        >
          <p>{strings.PanelContent}</p>
          <p>{token}</p>
        </Panel>
      </div>
    );
  }

  //#endregion

  //#region Helpers

  _onChange = (_ev: React.FormEvent<HTMLInputElement>, newValue?: string) => {
    this.setState({ projectRef: newValue || "" });
  };

  dismissPanel = () => {
    this.setState({ openSSOToken: false });
  };

  // _getProjectRef = async () => {
  //   await Word.run(async (context) => {
  //     let properties = context.document.properties.customProperties;
  //     properties.load("key,type,value");

  //     await context.sync();

  //     var projectRef: string;
  //     for (var i = 0; i < properties.items.length; i++) {
  //       if (properties.items[i].key === "ProjectRef") {
  //         projectRef = properties.items[i].value;
  //       }
  //     }

  //     this.setState({
  //       projectRef: projectRef,
  //     });
  //   });
  // };

  updateProjectRef = async () => {
    this._wordServices.updateProjectRef;

    this.setState({
      success: true,
    });

    setTimeout(() => {
      this.setState({
        success: false,
      });
    }, 2000);
  };

  //#endregion
}
