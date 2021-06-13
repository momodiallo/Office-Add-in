import * as React from "react";
import Home from "../outlook/Home";
import Login from "../Login";

import "../../../../assets/icon-16.png";
import "../../../../assets/icon-32.png";
import "../../../../assets/icon-80.png";

export interface OutlookProps {
  title: string;
  token: string;
  isUserLoggedIn: boolean;
}

export default class Outlook extends React.Component<OutlookProps> {
  render() {
    if (this.props.isUserLoggedIn) {
      return <Home token={this.props.token} />;
    } else {
      return <Login title={this.props.title} />;
    }
  }
}
