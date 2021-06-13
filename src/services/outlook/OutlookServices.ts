/* global Office */

import { IOutlookServices } from "./IOutlookServices";

export default class OutlookServices implements IOutlookServices {
  public getCurrentEmailItem() {
    try {
      let currentItem = Office.context.mailbox.item;
      let from = currentItem.from.emailAddress;
      let to = currentItem.to.map((t) => t.emailAddress).join("; ");
      let cc = currentItem.cc.map((c) => c.emailAddress).join("; ");
      let subject = currentItem.subject;
      let attachments = currentItem.attachments.map((a) => a.name).join("; ");
      let result = {
        from: from,
        to: to,
        cc: cc,
        subject: subject,
        attachments: attachments,
      };
      return result;
    } catch (error) {
      return null;
    }
  }

  public getCurrentEmailBody() {
    return new Promise((resolve, reject) => {
      Office.context.mailbox.item.body.getAsync("text", {}, function callback(result) {
        if (result.status == Office.AsyncResultStatus.Succeeded) {
          let body = result.value;
          resolve(body);
        } else {
          reject(result.error);
        }
      });
    });
  }
}
