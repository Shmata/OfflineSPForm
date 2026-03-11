import Dexie, { type Table } from 'dexie';

import type { IIssueItem } from '../models';

export type OfflineOperation = 'insert' | 'update' | 'delete';

export type OfflineSubmissionStatus = 'pending' | 'processing';

export interface IOfflineSubmission {
  id?: number;
  clientId: string;
  dedupeKey: string;
  createdAt: number;
  updatedAt: number;
  webUrl: string;
  listGuid: string;
  status: OfflineSubmissionStatus;
  operation: OfflineOperation;
  itemId?: number;
  payload?: IIssueItem;
  attempts: number;
  lastError?: string;
}

function newClientId(): string {
  try {
    // Most modern browsers.
    return crypto.randomUUID();
  } catch {
    // Fallback.
    return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
  }
}

function getDedupeKey(args: {
  webUrl: string;
  listGuid: string;
  operation: OfflineOperation;
  itemId?: number;
  clientId: string;
}): string {
  if ((args.operation === 'update' || args.operation === 'delete') && typeof args.itemId === 'number') {
    return `${args.webUrl}|${args.listGuid}|${args.operation}|${args.itemId}`;
  }

  // For inserts (no itemId), de-dupe per client submission.
  return `${args.webUrl}|${args.listGuid}|${args.operation}|${args.clientId}`;
}

type MutableUnknownRecord = Record<string, unknown>;

class OfflineFormDb extends Dexie {
  public submissions!: Table<IOfflineSubmission, number>;

  public constructor() {
    super('OfflineSpFormDb');

    // v1 schema
    this.version(1).stores({
      submissions: '++id, createdAt, webUrl, listGuid, operation, itemId'
    });

    // v2 schema: add queue coordination fields + unique dedupeKey + compound index for efficient claiming.
    this.version(2)
      .stores({
        submissions:
          '++id, createdAt, updatedAt, webUrl, listGuid, status, operation, itemId, attempts, &dedupeKey, [webUrl+listGuid+status+createdAt]'
      })
      .upgrade(async (tx) => {
        const table = tx.table('submissions') as Table<IOfflineSubmission, number>;
        await table.toCollection().modify((s: unknown) => {
          const rec = s as MutableUnknownRecord;

          const clientId = (typeof rec.clientId === 'string' && rec.clientId) ? (rec.clientId as string) : newClientId();
          const operation = (rec.operation as OfflineOperation) || 'insert';
          const webUrl = typeof rec.webUrl === 'string' ? rec.webUrl : '';
          const listGuid = typeof rec.listGuid === 'string' ? rec.listGuid : '';
          const itemId = typeof rec.itemId === 'number' ? rec.itemId : undefined;

          rec.clientId = clientId;
          rec.status = (rec.status as OfflineSubmissionStatus) || 'pending';
          rec.updatedAt = (typeof rec.updatedAt === 'number' ? rec.updatedAt : (typeof rec.createdAt === 'number' ? rec.createdAt : Date.now()));
          rec.attempts = (typeof rec.attempts === 'number' ? rec.attempts : 0);
          rec.dedupeKey = (typeof rec.dedupeKey === 'string' && rec.dedupeKey)
            ? rec.dedupeKey
            : getDedupeKey({ webUrl, listGuid, operation, itemId, clientId });
        });
      });
  }
}

const db = new OfflineFormDb();

export class OfflineSubmissionStore {
  public async enqueue(args: {
    createdAt?: number;
    webUrl: string;
    listGuid: string;
    operation: OfflineOperation;
    itemId?: number;
    payload?: IIssueItem;
  }): Promise<number> {
    const now = Date.now();
    const createdAt = typeof args.createdAt === 'number' ? args.createdAt : now;
    const clientId = newClientId();
    const dedupeKey = getDedupeKey({
      webUrl: args.webUrl,
      listGuid: args.listGuid,
      operation: args.operation,
      itemId: args.itemId,
      clientId
    });

    return db.transaction('rw', db.submissions, async () => {
      const existing = await db.submissions.where('dedupeKey').equals(dedupeKey).first();

      if (existing?.id) {
        await db.submissions.update(existing.id, {
          // keep existing.clientId to remain stable for this queued item
          updatedAt: now,
          createdAt: existing.createdAt || createdAt,
          webUrl: args.webUrl,
          listGuid: args.listGuid,
          status: 'pending',
          operation: args.operation,
          itemId: args.itemId,
          payload: args.payload,
          lastError: undefined
        });
        return existing.id;
      }

      return db.submissions.add({
        clientId,
        dedupeKey,
        createdAt,
        updatedAt: now,
        webUrl: args.webUrl,
        listGuid: args.listGuid,
        status: 'pending',
        operation: args.operation,
        itemId: args.itemId,
        payload: args.payload,
        attempts: 0
      });
    });
  }

  public async listPending(webUrl: string, listGuid: string): Promise<IOfflineSubmission[]> {
    return db.submissions
      .where({ webUrl, listGuid })
      .filter((s) => s.status === 'pending')
      .sortBy('createdAt');
  }

  public async takeNextPending(webUrl: string, listGuid: string): Promise<IOfflineSubmission | undefined> {
    return db.transaction('rw', db.submissions, async () => {
      const first = await db.submissions
        .where('[webUrl+listGuid+status+createdAt]')
        .between([webUrl, listGuid, 'pending', Dexie.minKey], [webUrl, listGuid, 'pending', Dexie.maxKey])
        .first();

      if (!first?.id) return undefined;

      const now = Date.now();
      await db.submissions.update(first.id, { status: 'processing', updatedAt: now });
      return { ...first, status: 'processing', updatedAt: now };
    });
  }

  public async markPending(id: number, errorMessage?: string): Promise<void> {
    await db.transaction('rw', db.submissions, async () => {
      const current = await db.submissions.get(id);
      const nextAttempts = (current?.attempts ?? 0) + 1;
      await db.submissions.update(id, {
        status: 'pending',
        updatedAt: Date.now(),
        attempts: nextAttempts,
        lastError: errorMessage
      });
    });
  }

  public async remove(id: number): Promise<void> {
    await db.submissions.delete(id);
  }
}
