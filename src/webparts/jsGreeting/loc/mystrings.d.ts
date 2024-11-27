declare interface IJsGreetingWebPartStrings {
  PropertyPaneGreeting: string;
  BasicGroupName: string;
  GreetingFieldLabel: string;
  BorderFieldLabel: number;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
}

declare module 'JsGreetingWebPartStrings' {
  const strings: IJsGreetingWebPartStrings;
  export = strings;
}
