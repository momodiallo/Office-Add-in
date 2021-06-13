import * as React from "react";
import { MessageBar, MessageBarType, PrimaryButton } from "office-ui-fabric-react";
import Header from "./Header";
import { getLocaleStrings } from "../loc/getLocaleStrings";
import { IGetAccessToken } from "../../services/IGetAccessToken";
import GetAccessToken from "../../services/GetAccessToken";
import Home from "./word/Home";

export interface ILoginProps {
  title: string;
}
export interface ILoginState {
  results: any;
  showMessageBar: boolean;
  isUserLoggedIn: boolean;
}

export default class Login extends React.Component<ILoginProps, ILoginState> {
  private readonly _getAccessTokenService: IGetAccessToken;

  constructor(props, context) {
    super(props, context);
    this.state = {
      results: {},
      showMessageBar: false,
      isUserLoggedIn: false,
    };

    this._getAccessTokenService = new GetAccessToken();
  }

  render() {
    const strings = getLocaleStrings();
    if (!this.state.isUserLoggedIn) {
      return (
        <div>
          <div className="ms-welcome">
            <Header logo="assets/lexor_logo.png" title={this.props.title} message="Lexor" />
            <h1>{strings.Heading}</h1>
            <p className="content">{strings.Introduction}</p>
            <PrimaryButton className="ms-welcome__action" onClick={this._login}>
              {strings.OpenAddinButton}
            </PrimaryButton>
          </div>
          {this.state.showMessageBar && (
            <MessageBar
              className="messageBar"
              onDismiss={() => this.setState({ showMessageBar: false })}
              messageBarType={MessageBarType.error}
              isMultiline={false}
            >
              {this.state.results?.errorMessage?.message}
            </MessageBar>
          )}
        </div>
      );
    } else {
      return <Home token={this.state.results.token} />;
    }
  }

  _login = async () => {
    let results = await this._getAccessTokenService.getAcessToken();
    this.setState({
      results: results,
      showMessageBar: results?.error,
      isUserLoggedIn: results?.error ? false : true,
    });
  };
}
