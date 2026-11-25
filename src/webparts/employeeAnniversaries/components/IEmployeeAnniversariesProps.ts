import type { SPHttpClient } from '@microsoft/sp-http';

export interface IEmployeeAnniversariesProps {
  title: string;
  webUrl: string;
  spHttpClient: SPHttpClient;
  listId: string;
  employeeField: string;
  startDateField: string;
  daysBefore: number;
  daysAfter: number;
  showPhoto: boolean;
  showName: boolean;
  showRole: boolean;
  showDepartment: boolean;
  showYears: boolean;
}
