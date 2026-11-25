import * as React from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import { Persona, PersonaSize, Spinner, MessageBar, MessageBarType } from '@fluentui/react';
import * as strings from 'EmployeeAnniversariesWebPartStrings';
import styles from './EmployeeAnniversaries.module.scss';
import type { IEmployeeAnniversariesProps } from './IEmployeeAnniversariesProps';
import { findAnniversary } from './anniversaryUtils';

interface IEmployee {
  key: string;
  name: string;
  email: string;
  jobTitle: string;
  department: string;
  date: Date;
  offsetDays: number;
  years: number;
}

interface IState {
  loading: boolean;
  error?: string;
  items: IEmployee[];
}

interface IListItem {
  Id: number;
  [key: string]: unknown;
}

interface IPerson {
  Id: number;
  Title?: string;
  EMail?: string;
  JobTitle?: string;
  Department?: string;
}

export default class EmployeeAnniversaries extends React.Component<IEmployeeAnniversariesProps, IState> {
  private _requestId: number = 0;

  public constructor(props: IEmployeeAnniversariesProps) {
    super(props);
    this.state = { loading: false, items: [] };
  }

  public componentDidMount(): void {
    void this._load();
  }

  public componentDidUpdate(prev: IEmployeeAnniversariesProps): void {
    const p = this.props;
    if (
      p.listId !== prev.listId ||
      p.employeeField !== prev.employeeField ||
      p.startDateField !== prev.startDateField ||
      p.daysBefore !== prev.daysBefore ||
      p.daysAfter !== prev.daysAfter
    ) {
      void this._load();
    }
  }

  private _isConfigured(): boolean {
    const { listId, employeeField, startDateField } = this.props;
    return !!(listId && employeeField && startDateField);
  }

  private async _load(): Promise<void> {
    if (!this._isConfigured()) {
      this.setState({ loading: false, error: undefined, items: [] });
      return;
    }

    const requestId = ++this._requestId;
    this.setState({ loading: true, error: undefined });

    const { webUrl, spHttpClient, listId, employeeField, startDateField, daysBefore, daysAfter } = this.props;
    try {
      const select = [
        'Id',
        startDateField,
        `${employeeField}/Id`,
        `${employeeField}/Title`,
        `${employeeField}/EMail`,
        `${employeeField}/JobTitle`,
        `${employeeField}/Department`
      ].join(',');
      let url: string | undefined =
        `${webUrl}/_api/web/lists(guid'${listId}')/items?$select=${select}&$expand=${employeeField}&$top=5000`;

      const today = new Date();
      const results: IEmployee[] = [];

      while (url) {
        const response = await spHttpClient.get(url, SPHttpClient.configurations.v1);
        if (!response.ok) {
          throw new Error(`${response.status} ${response.statusText}`);
        }
        const json: { value: IListItem[]; '@odata.nextLink'?: string } = await response.json();

        for (const item of json.value) {
          const rawDate = item[startDateField] as string | undefined;
          const person = item[employeeField] as IPerson | undefined;
          if (!rawDate || !person) {
            continue;
          }
          const start = new Date(rawDate);
          if (isNaN(start.getTime())) {
            continue;
          }
          // Date-only values arrive as UTC midnight; use UTC parts to avoid a day shift.
          const localStart = new Date(start.getUTCFullYear(), start.getUTCMonth(), start.getUTCDate());
          const match = findAnniversary(localStart, today, daysBefore, daysAfter);
          if (!match) {
            continue;
          }
          results.push({
            key: `${item.Id}`,
            name: person.Title || '',
            email: person.EMail || '',
            jobTitle: person.JobTitle || '',
            department: person.Department || '',
            date: match.date,
            offsetDays: match.offsetDays,
            years: match.years
          });
        }
        url = json['@odata.nextLink'];
      }

      results.sort((a, b) => a.date.getTime() - b.date.getTime() || a.name.localeCompare(b.name));
      if (requestId === this._requestId) {
        this.setState({ loading: false, items: results });
      }
    } catch (e) {
      if (requestId === this._requestId) {
        this.setState({ loading: false, items: [], error: (e as Error).message });
      }
    }
  }

  private _whenLabel(offset: number): string {
    if (offset === 0) return strings.Today;
    if (offset === 1) return strings.Tomorrow;
    if (offset === -1) return strings.Yesterday;
    return offset > 0
      ? strings.InDays.replace('{0}', `${offset}`)
      : strings.DaysAgo.replace('{0}', `${-offset}`);
  }

  private _renderBody(): React.ReactNode {
    const { loading, error, items } = this.state;
    const { showPhoto, showName, showRole, showDepartment, showYears, webUrl } = this.props;

    if (!this._isConfigured()) {
      return <MessageBar>{strings.NotConfigured}</MessageBar>;
    }
    if (loading) {
      return <Spinner label={strings.Loading} />;
    }
    if (error) {
      return <MessageBar messageBarType={MessageBarType.error}>{strings.ErrorPrefix}{error}</MessageBar>;
    }
    if (items.length === 0) {
      return <MessageBar>{strings.NoAnniversaries}</MessageBar>;
    }

    return (
      <ul className={styles.list}>
        {items.map(e => {
          const secondary = [showRole ? e.jobTitle : '', showDepartment ? e.department : '']
            .filter(s => !!s)
            .join(' · ');
          const photo = e.email
            ? `${webUrl}/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(e.email)}`
            : undefined;
          return (
            <li key={e.key} className={`${styles.item} ${e.offsetDays === 0 ? styles.today : ''}`}>
              <div className={styles.person}>
                <Persona
                  text={showName ? e.name : ''}
                  secondaryText={secondary}
                  imageUrl={showPhoto ? photo : undefined}
                  imageAlt={e.name}
                  hidePersonaDetails={!showName && !secondary}
                  size={PersonaSize.size48}
                />
              </div>
              <div className={styles.when}>
                <div className={styles.whenLabel}>{this._whenLabel(e.offsetDays)}</div>
                <div className={styles.date}>
                  {e.date.toLocaleDateString(undefined, { day: 'numeric', month: 'short' })}
                  {showYears && ` · ${e.years} ${e.years === 1 ? strings.Year : strings.Years}`}
                </div>
              </div>
            </li>
          );
        })}
      </ul>
    );
  }

  public render(): React.ReactElement<IEmployeeAnniversariesProps> {
    return (
      <section className={styles.employeeAnniversaries}>
        {this.props.title && <h2 className={styles.title}>{this.props.title}</h2>}
        {this._renderBody()}
      </section>
    );
  }
}
