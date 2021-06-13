/* global Office */
export interface IOutlookServices {
  getCurrentEmailItem(): {
    from: string;
    to: string;
    cc: string;
    subject: Office.Subject & string;
    attachments: string;
  };

  getCurrentEmailBody(): Promise<unknown>;
}
