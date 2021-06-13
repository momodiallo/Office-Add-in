import * as React from "react";
import * as ReactDOM from "react-dom";
import "office-ui-fabric-react/dist/css/fabric.min.css";
import Word from "./components/word/Word";
import Outlook from "./components/outlook/Outlook";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "office-ui-fabric-react/lib/Icons";

/* global document, Office, OfficeRuntime */
initializeIcons();
const title = "Lexor Add-in";
let token: string;
let isUserLoggedIn: boolean;

const render = (Component) => {
  ReactDOM.render(
    <AppContainer>
      <Component title={title} token={token} isUserLoggedIn={isUserLoggedIn} />
    </AppContainer>,
    document.getElementById("container")
  );
};

/* Render application after Office initializes */
Office.initialize = function () {};

Office.onReady(async (info) => {
  token = await _getAcessToken();
  isUserLoggedIn = token ? true : false;

  switch (info.host) {
    case Office.HostType.Word:
      render(Word);
      break;
    case Office.HostType.Outlook:
      render(Outlook);
      break;
    default:
      break;
  }
});

async function _getAcessToken() {
  let bootstrapToken: string;
  try {
    bootstrapToken = await OfficeRuntime.auth.getAccessToken();
    return bootstrapToken;
  } catch (error) {
    return null;
  }
}
