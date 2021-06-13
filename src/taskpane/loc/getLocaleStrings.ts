import strings_EN from "./en-us";
import strings_FR from "./fr-fr";
import strings_NL from "./nl-nl";

/* global Office*/

export function getLocaleStrings() {
  let userLanguage = Office.context.displayLanguage;
  let localeStrings: any;

  // Get the resource strings that match the language.
  switch (userLanguage.toLocaleLowerCase()) {
    case "en-us":
      localeStrings = strings_EN;
      break;
    case "fr-fr":
      localeStrings = strings_FR;
      break;
    case "nl-nl":
      localeStrings = strings_NL;
      break;
    default:
      localeStrings = strings_EN;
      break;
  }

  return localeStrings;
}
