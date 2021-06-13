/* global OfficeRuntime */

import { IGetAccessToken } from "./IGetAccessToken";

export default class GetAccessToken implements IGetAccessToken {
  public async getAcessToken() {
    try {
      var response = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true });
      return {
        isUserLoggedIn: true,
        token: response,
      };
    } catch (error) {
      return {
        error: true,
        errorMessage: error,
      };
    }
  }
}
