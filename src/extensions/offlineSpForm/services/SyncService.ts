import { Log } from '@microsoft/sp-core-library';
import type { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';

import type { IIssueItem } from '../models';
import { NetworkService } from './NetworkService';
import { OfflineSubmissionStore, type IOfflineSubmission } from './OfflineSubmissionStore';
import { Services } from './Services';

const LOG_SOURCE = 'OfflineSpForm.SyncService';

export class SyncService {
  private static _inFlightByKey: Map<string, Promise<void>> = new Map();

  private readonly webUrl: string;
  private readonly services: Services;
  private readonly store: OfflineSubmissionStore;

  public constructor(
    private readonly context: FormCustomizerContext,
    private readonly listGuid: string
  ) {
    this.webUrl = context.pageContext.web.absoluteUrl;
    this.services = new Services(context, listGuid);
    this.store = new OfflineSubmissionStore();
  }

  public async syncPending(maxItems: number = 50): Promise<void> {
    const state = NetworkService.getCurrentState();
    if (!state.online || state.isWeak) return;

    const key = `${this.webUrl}|${this.listGuid}`;
    const existing = SyncService._inFlightByKey.get(key);
    if (existing) return;

    const run = this._syncPendingInternal(maxItems)
      .then(() => {
        SyncService._inFlightByKey.delete(key);
      })
      .catch((e) => {
        SyncService._inFlightByKey.delete(key);
        throw e;
      });

    SyncService._inFlightByKey.set(key, run);
    return run;
  }

  private async _syncPendingInternal(maxItems: number): Promise<void> {
    let processed = 0;

    while (processed < maxItems) {
      const state = NetworkService.getCurrentState();
      if (!state.online || state.isWeak) return;

      const next = await this.store.takeNextPending(this.webUrl, this.listGuid);
      if (!next) return;

      try {
        await this._applySubmission(next);
        if (typeof next.id === 'number') {
          await this.store.remove(next.id);
        }
        processed++;
      } catch (e) {
        const msg = e instanceof Error ? e.message : String(e);
        if (typeof next.id === 'number') {
          await this.store.markPending(next.id, msg);
        }
        Log.error(LOG_SOURCE, e instanceof Error ? e : new Error(msg));

        // Stop this run to avoid hot-looping on a bad entry.
        return;
      }
    }
  }

  private async _applySubmission(submission: IOfflineSubmission): Promise<void> {
    if (submission.operation === 'insert') {
      const payload = submission.payload as IIssueItem | undefined;
      if (!payload) throw new Error('Missing payload for insert.');
      await this.services.insertItem(payload);
      return;
    }

    if (submission.operation === 'update') {
      if (typeof submission.itemId !== 'number') throw new Error('Missing itemId for update.');
      const payload = submission.payload as IIssueItem | undefined;
      if (!payload) throw new Error('Missing payload for update.');
      await this.services.updateItem(submission.itemId, payload);
      return;
    }

    if (submission.operation === 'delete') {
      if (typeof submission.itemId !== 'number') throw new Error('Missing itemId for delete.');
      await this.services.deleteItem(submission.itemId);
      return;
    }

    // Should be unreachable due to the OfflineOperation union.
    throw new Error(`Unknown operation: ${String(submission.operation)}`);
  }
}
