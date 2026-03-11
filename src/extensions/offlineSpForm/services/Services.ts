import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';

import type {
  IGetItemsOptions,
  IIssueItem,
  ISharePointLookupValue,
  ISharePointPersonValue,
  ISpIdTitleRow,
  ISpIssueRow,
  ISpListResponse,
  ISpSiteUserRow
} from '../models';

export class Services {
  private readonly webUrl: string;

  private static readonly _getHeaders: Record<string, string> = {
    Accept: 'application/json;odata=nometadata',
    'odata-version': ''
  };

  private static readonly _writeHeaders: Record<string, string> = {
    Accept: 'application/json;odata=nometadata',
    'Content-Type': 'application/json;odata=nometadata',
    'odata-version': ''
  };

  public constructor(
    private readonly context: FormCustomizerContext,
    private readonly listGuid: string
  ) {
    this.webUrl = context.pageContext.web.absoluteUrl;
  }

  public async searchUsers(queryText: string, top: number = 10): Promise<ISharePointPersonValue[]> {
    const q = (queryText || '').trim();
    if (!q) return [];

    const escaped = q.replace(/'/g, "''");
    const filter = `startswith(Title,'${escaped}') or startswith(Email,'${escaped}')`;
    const select = ['Id', 'Title', 'Email'].join(',');
    const queryParts: string[] = [
      `$select=${encodeURIComponent(select)}`,
      `$top=${top}`,
      `$filter=${encodeURIComponent(filter)}`
    ];

    const url = `${this.webUrl}/_api/web/siteusers?${queryParts.join('&')}`;

    const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
      headers: {
        ...Services._getHeaders
      }
    });

    await this._throwIfNotOk(response);
    const data = (await response.json()) as ISpListResponse<ISpSiteUserRow>;

    return (data.value || []).map((u) => ({
      Id: u.Id,
      Title: u.Title,
      EMail: u.Email
    }));
  }

  public async getUserById(userId: number): Promise<ISharePointPersonValue | undefined> {
    const url = `${this.webUrl}/_api/web/getuserbyid(${userId})?$select=Id,Title,Email`;

    const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
      headers: {
        ...Services._getHeaders
      }
    });

    if (!response.ok) return undefined;
    const u = (await response.json()) as ISpSiteUserRow;

    return {
      Id: u.Id,
      Title: u.Title,
      EMail: u.Email
    };
  }

  public async getItems(options: IGetItemsOptions = {}): Promise<IIssueItem[]> {
    const select = [
      'Id',
      'Title',
      'Description',
      'Priority',
      'Status',
      'Assignedto0/Id',
      'Assignedto0/Title',
      'Assignedto0/EMail',
      'RelatedIssue/Id',
      'RelatedIssue/Title'
    ].join(',');

    const expand = ['Assignedto0', 'RelatedIssue'].join(',');

    const queryParts: string[] = [`$select=${encodeURIComponent(select)}`, `$expand=${encodeURIComponent(expand)}`];
    if (options.top) queryParts.push(`$top=${options.top}`);
    if (options.filter) queryParts.push(`$filter=${encodeURIComponent(options.filter)}`);
    if (options.orderBy) queryParts.push(`$orderby=${encodeURIComponent(options.orderBy)}`);

    const url = `${this.webUrl}/_api/web/lists(guid'${this.listGuid}')/items?${queryParts.join('&')}`;

    const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
      headers: {
        ...Services._getHeaders
      }
    });

    await this._throwIfNotOk(response);
    const data = (await response.json()) as ISpListResponse<ISpIssueRow>;

    return (data.value || []).map((row) => this._fromSpRow(row));
  }

  public async getItemTitles(listGuid: string = this.listGuid, top: number = 200): Promise<ISharePointLookupValue[]> {
    const select = ['Id', 'Title'].join(',');
    const queryParts: string[] = [
      `$select=${encodeURIComponent(select)}`,
      `$top=${top}`,
      `$orderby=${encodeURIComponent('Title asc')}`
    ];

    const url = `${this.webUrl}/_api/web/lists(guid'${listGuid}')/items?${queryParts.join('&')}`;
    const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
      headers: {
        ...Services._getHeaders
      }
    });

    await this._throwIfNotOk(response);
    const data = (await response.json()) as ISpListResponse<ISpIdTitleRow>;
    return (data.value || []).map((r) => ({
      Id: r.Id,
      Title: r.Title
    }));
  }

  public async insertItem(item: IIssueItem): Promise<IIssueItem> {
    const url = `${this.webUrl}/_api/web/lists(guid'${this.listGuid}')/items`;

    const response = await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        ...Services._writeHeaders
      },
      body: JSON.stringify(this._toSpFields(item))
    });

    await this._throwIfNotOk(response);
    const created = (await response.json()) as ISpIssueRow;
    return this._fromSpRow(created);
  }

  public async updateItem(itemId: number, item: Partial<IIssueItem>): Promise<void> {
    const url = `${this.webUrl}/_api/web/lists(guid'${this.listGuid}')/items(${itemId})`;

    const response = await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        ...Services._writeHeaders,
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      },
      body: JSON.stringify(this._toSpFields(item))
    });

    await this._throwIfNotOk(response);
  }

  public async deleteItem(itemId: number): Promise<void> {
    const url = `${this.webUrl}/_api/web/lists(guid'${this.listGuid}')/items(${itemId})`;

    const response = await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        ...Services._writeHeaders,
        'IF-MATCH': '*',
        'X-HTTP-Method': 'DELETE'
      }
    });

    await this._throwIfNotOk(response);
  }

  private _toSpFields(item: Partial<IIssueItem>): Record<string, unknown> {
    const fields: Record<string, unknown> = {};

    if (item.Title !== undefined) fields.Title = item.Title;
    if (item.Description !== undefined) fields.Description = item.Description;
    if (item.Priority !== undefined) fields.Priority = item.Priority;
    if (item.Status !== undefined) fields.Status = item.Status;

    if (item.Assignedto0 !== undefined) {
      const personId = this._getId(item.Assignedto0);
      fields.Assignedto0Id = personId ?? null;
    }

    if (item.RelatedIssue !== undefined) {
      const lookupId = this._getId(item.RelatedIssue);
      fields.RelatedIssueId = lookupId ?? null;
    }

    return fields;
  }

  private _fromSpRow(row: ISpIssueRow): IIssueItem {
    const assigned: ISharePointPersonValue | undefined = row.Assignedto0
      ? {
          Id: row.Assignedto0.Id,
          Title: row.Assignedto0.Title,
          EMail: row.Assignedto0.EMail
        }
      : undefined;

    const related: ISharePointLookupValue | undefined = row.RelatedIssue
      ? {
          Id: row.RelatedIssue.Id,
          Title: row.RelatedIssue.Title
        }
      : undefined;

    return {
      Id: row.Id,
      Title: row.Title ?? '',
      Description: row.Description,
      Priority: row.Priority,
      Status: row.Status,
      Assignedto0: assigned,
      RelatedIssue: related
    };
  }

  private _getId(value: unknown): number | undefined {
    if (value === null || value === undefined) return undefined;
    if (typeof value === 'number') return value;

    const maybe = value as { Id?: number };
    if (typeof maybe.Id === 'number') return maybe.Id;

    return undefined;
  }

  private async _throwIfNotOk(response: SPHttpClientResponse): Promise<void> {
    if (response.ok) return;

    let message = `${response.status} ${response.statusText}`;
    try {
      const text = await response.text();
      if (text) message = `${message}: ${text}`;
    } catch (e: unknown) {
      console.error(e);
    }

    throw new Error(message);
  }
}
