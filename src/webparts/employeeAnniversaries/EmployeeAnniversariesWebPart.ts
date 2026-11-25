import * as React from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  type IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPHttpClient } from '@microsoft/sp-http';

import * as strings from 'EmployeeAnniversariesWebPartStrings';
import EmployeeAnniversaries from './components/EmployeeAnniversaries';
import { IEmployeeAnniversariesProps } from './components/IEmployeeAnniversariesProps';

export interface IEmployeeAnniversariesWebPartProps {
  title: string;
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

interface ISPField {
  InternalName: string;
  Title: string;
}

export default class EmployeeAnniversariesWebPart extends BaseClientSideWebPart<IEmployeeAnniversariesWebPartProps> {

  private _root: Root | undefined;
  private _lists: IPropertyPaneDropdownOption[] = [];
  private _userFields: IPropertyPaneDropdownOption[] = [];
  private _dateFields: IPropertyPaneDropdownOption[] = [];
  private _loadingOptions: boolean = false;

  public render(): void {
    const element: React.ReactElement<IEmployeeAnniversariesProps> = React.createElement(
      EmployeeAnniversaries,
      {
        title: this.properties.title,
        webUrl: this.context.pageContext.web.absoluteUrl,
        spHttpClient: this.context.spHttpClient,
        listId: this.properties.listId,
        employeeField: this.properties.employeeField,
        startDateField: this.properties.startDateField,
        daysBefore: this.properties.daysBefore,
        daysAfter: this.properties.daysAfter,
        showPhoto: this.properties.showPhoto,
        showName: this.properties.showName,
        showRole: this.properties.showRole,
        showDepartment: this.properties.showDepartment,
        showYears: this.properties.showYears
      }
    );

    if (!this._root) {
      this._root = createRoot(this.domElement);
    }
    this._root.render(element);
  }

  protected onInit(): Promise<void> {
    const p = this.properties;
    p.title = p.title ?? strings.DefaultTitle;
    p.daysBefore = p.daysBefore ?? 7;
    p.daysAfter = p.daysAfter ?? 7;
    p.showPhoto = p.showPhoto ?? true;
    p.showName = p.showName ?? true;
    p.showRole = p.showRole ?? true;
    p.showDepartment = p.showDepartment ?? true;
    p.showYears = p.showYears ?? true;
    return Promise.resolve();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }
    const { semanticColors } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
    }
  }

  protected onDispose(): void {
    this._root?.unmount();
    this._root = undefined;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private async _get<T>(path: string): Promise<T[]> {
    const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/${path}`;
    const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!response.ok) {
      throw new Error(`${response.status} ${response.statusText}`);
    }
    return (await response.json()).value as T[];
  }

  private async _loadLists(): Promise<void> {
    const lists = await this._get<{ Id: string; Title: string }>(
      'lists?$select=Id,Title&$filter=Hidden eq false and BaseTemplate eq 100&$orderby=Title'
    );
    this._lists = lists.map(l => ({ key: l.Id, text: l.Title }));
  }

  private async _loadFields(): Promise<void> {
    this._userFields = [];
    this._dateFields = [];
    if (!this.properties.listId) {
      return;
    }
    const base = `lists(guid'${this.properties.listId}')/fields?$select=InternalName,Title&$filter=Hidden eq false and ReadOnlyField eq false`;
    const [users, dates] = await Promise.all([
      this._get<ISPField>(`${base} and TypeAsString eq 'User'`),
      this._get<ISPField>(`${base} and TypeAsString eq 'DateTime'`)
    ]);
    this._userFields = users.map(f => ({ key: f.InternalName, text: f.Title }));
    this._dateFields = dates.map(f => ({ key: f.InternalName, text: f.Title }));
  }

  protected onPropertyPaneConfigurationStart(): void {
    this._loadingOptions = true;
    this.context.propertyPane.refresh();
    Promise.all([this._loadLists(), this._loadFields()])
      .catch(() => { /* leave dropdowns empty */ })
      .then(() => {
        this._loadingOptions = false;
        this.context.propertyPane.refresh();
      })
      .catch(() => { /* ignore */ });
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if (propertyPath === 'listId' && oldValue !== newValue) {
      this.properties.employeeField = '';
      this.properties.startDateField = '';
      this._loadingOptions = true;
      this.context.propertyPane.refresh();
      this._loadFields()
        .catch(() => { /* leave empty */ })
        .then(() => {
          this._loadingOptions = false;
          this.context.propertyPane.refresh();
          this.render();
        })
        .catch(() => { /* ignore */ });
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const disabled = this._loadingOptions;
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.DataGroupName,
              groupFields: [
                PropertyPaneTextField('title', { label: strings.TitleLabel }),
                PropertyPaneDropdown('listId', {
                  label: strings.ListLabel,
                  options: this._lists,
                  disabled
                }),
                PropertyPaneDropdown('employeeField', {
                  label: strings.EmployeeFieldLabel,
                  options: this._userFields,
                  disabled: disabled || !this.properties.listId
                }),
                PropertyPaneDropdown('startDateField', {
                  label: strings.StartDateFieldLabel,
                  options: this._dateFields,
                  disabled: disabled || !this.properties.listId
                }),
                PropertyPaneSlider('daysBefore', {
                  label: strings.DaysBeforeLabel,
                  min: 0, max: 60, step: 1, showValue: true
                }),
                PropertyPaneSlider('daysAfter', {
                  label: strings.DaysAfterLabel,
                  min: 0, max: 60, step: 1, showValue: true
                })
              ]
            },
            {
              groupName: strings.DisplayGroupName,
              groupFields: [
                PropertyPaneToggle('showPhoto', { label: strings.ShowPhotoLabel }),
                PropertyPaneToggle('showName', { label: strings.ShowNameLabel }),
                PropertyPaneToggle('showRole', { label: strings.ShowRoleLabel }),
                PropertyPaneToggle('showDepartment', { label: strings.ShowDepartmentLabel }),
                PropertyPaneToggle('showYears', { label: strings.ShowYearsLabel })
              ]
            }
          ]
        }
      ]
    };
  }
}
