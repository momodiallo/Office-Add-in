export interface IGetAccessToken {
  getAcessToken(): Promise<
    | {
        isUserLoggedIn: boolean;
        token: string;
        error?: undefined;
        errorMessage?: undefined;
      }
    | {
        error: boolean;
        errorMessage: any;
        isUserLoggedIn?: undefined;
        token?: undefined;
      }
  >;
}
