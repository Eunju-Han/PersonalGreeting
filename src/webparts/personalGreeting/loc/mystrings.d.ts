declare interface IPersonalGreetingWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  SecondaryGroupName: string;
  TertiaryGroupName: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'PersonalGreetingWebPartStrings' {
  const strings: IPersonalGreetingWebPartStrings;
  export = strings;
}
