import { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';

export interface IContextItemRow {
  Id?: number;
  Title?: string;
  Description?: string;
  Priority?: string;
  Status?: string;
  Assignedto0Id?: number;
  RelatedIssueId?: number;
}

export type IContextWithItem = FormCustomizerContext & {
  item?: IContextItemRow;
  itemId?: number;
};
