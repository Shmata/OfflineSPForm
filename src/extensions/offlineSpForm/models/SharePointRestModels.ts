export interface IGetItemsOptions {
  top?: number;
  filter?: string;
  orderBy?: string;
}

export interface ISpListResponse<T> {
  value: T[];
}

export interface ISpPersonRow {
  Id: number;
  Title?: string;
  EMail?: string;
}

export interface ISpLookupRow {
  Id: number;
  Title?: string;
}

export interface ISpIssueRow {
  Id?: number;
  Title?: string;
  Description?: string;
  Priority?: string;
  Status?: string;
  Assignedto0?: ISpPersonRow;
  RelatedIssue?: ISpLookupRow;
}

export interface ISpSiteUserRow {
  Id: number;
  Title?: string;
  Email?: string;
  LoginName?: string;
}

export interface ISpIdTitleRow {
  Id: number;
  Title?: string;
}
