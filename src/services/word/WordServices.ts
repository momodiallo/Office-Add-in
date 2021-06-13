/* global Word */

import { IWordServices } from "./IWordServices";

export default class WordServices implements IWordServices {
  public async updateProjectRef(projectRef) {
    await Word.run(async (context) => {
      context.document.properties.customProperties.add("ProjectRef", projectRef);
      await context.sync();
    });
  }

  public async getProjectRef() {
    await Word.run(async (context) => {
      let properties = context.document.properties.customProperties;
      properties.load("key,type,value");

      await context.sync();

      var projectRef: string;
      for (var i = 0; i < properties.items.length; i++) {
        if (properties.items[i].key === "ProjectRef") {
          projectRef = properties.items[i].value;
        }
      }
      return projectRef;
    });
  }
}
