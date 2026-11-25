declare interface IEmployeeAnniversariesWebPartStrings {
  PropertyPaneDescription: string;
  DataGroupName: string;
  DisplayGroupName: string;
  TitleLabel: string;
  ListLabel: string;
  EmployeeFieldLabel: string;
  StartDateFieldLabel: string;
  DaysBeforeLabel: string;
  DaysAfterLabel: string;
  ShowPhotoLabel: string;
  ShowNameLabel: string;
  ShowRoleLabel: string;
  ShowDepartmentLabel: string;
  ShowYearsLabel: string;
  LoadingOptions: string;
  NotConfigured: string;
  NoAnniversaries: string;
  Loading: string;
  ErrorPrefix: string;
  Today: string;
  Tomorrow: string;
  Yesterday: string;
  InDays: string;
  DaysAgo: string;
  Year: string;
  Years: string;
  DefaultTitle: string;
}

declare module 'EmployeeAnniversariesWebPartStrings' {
  const strings: IEmployeeAnniversariesWebPartStrings;
  export = strings;
}
