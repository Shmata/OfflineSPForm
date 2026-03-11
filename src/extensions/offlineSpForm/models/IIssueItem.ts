export interface ISharePointLookupValue {
  Id: number;
  Title?: string;
}

export interface ISharePointPersonValue {
  Id: number;
  Title?: string;
  EMail?: string;
}

export interface IIssueItem {
  Id?: number;
  Title: string;
  Description?: string;
  Priority?: string;
  Status?: string;
  Assignedto0?: ISharePointPersonValue;
  RelatedIssue?: ISharePointLookupValue;
}
