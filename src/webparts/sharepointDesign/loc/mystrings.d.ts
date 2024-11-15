declare interface ISharepointDesignWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'SharepointDesignWebPartStrings' {
  const strings: ISharepointDesignWebPartStrings;
  export = strings;
}
