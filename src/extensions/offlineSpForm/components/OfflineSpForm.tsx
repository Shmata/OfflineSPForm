import * as React from 'react';
import { useEffect, useState, useMemo, useCallback } from 'react';
import { Log, FormDisplayMode } from '@microsoft/sp-core-library';
import { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';
import { DefaultButton, Dropdown, getTheme, IDropdownOption, Label, MessageBar, MessageBarType, PrimaryButton, Stack, TextField } from '@fluentui/react';
import { NormalPeoplePicker } from '@fluentui/react/lib/Pickers';
import type { IPersonaProps } from '@fluentui/react/lib/Persona';

import styles from './OfflineSpForm.module.scss';

import { IContextWithItem, IIssueItem } from '../models';
import { NetworkService, OfflineSubmissionStore, Services, SyncService } from '../services';

export interface IOfflineSpFormProps {
  context: FormCustomizerContext;
  displayMode: FormDisplayMode;
  onSave: () => void;
  onClose: () => void;
}

const LOG_SOURCE: string = 'OfflineSpForm';

const LIST_GUID: string = 'Replace-with-your-list-guid'; // TODO: replace with your list's GUID
const RELATED_ISSUE_LIST_GUID: string = LIST_GUID;

const PRIORITY_CHOICES: string[] = ['Critical', 'High', 'Normal', 'Low'];
const STATUS_CHOICES: string[] = ['New', 'In Progress', 'Completed', 'Blocked', 'Duplicate', 'By design', 'Won\'t fix'];

const PRIORITY_OPTIONS: IDropdownOption[] = PRIORITY_CHOICES.map((c) => ({ key: c, text: c }));
const STATUS_OPTIONS: IDropdownOption[] = STATUS_CHOICES.map((c) => ({ key: c, text: c }));

const OfflineSpForm: React.FC<IOfflineSpFormProps> = (props) => {

  const isReadOnly: boolean = props.displayMode === FormDisplayMode.Display;

  const [issue, setIssue] = useState<IIssueItem>({
    Title: '',
    Description: '',
    Priority: '',
    Status: ''
  });

  const [assignedPeople, setAssignedPeople] = useState<IPersonaProps[]>([]);
  const [relatedIssueId, setRelatedIssueId] = useState<number | undefined>(undefined);
  const [relatedIssueOptions, setRelatedIssueOptions] = useState<IDropdownOption[]>([]);
  const [isSaving, setIsSaving] = useState<boolean>(false);
  const [error, setError] = useState<string | undefined>(undefined);
  const [success, setSuccess] = useState<string | undefined>(undefined);

  const [network, setNetwork] = useState(() => NetworkService.getCurrentState());
  const isWeakNetwork: boolean = network.isWeak;
  const isStableOnline: boolean = network.online && !network.isWeak;

  const theme = useMemo(() => getTheme(), []);

  const services = useMemo(() => new Services(props.context, LIST_GUID), [props.context]);
  const offlineStore = useMemo(() => new OfflineSubmissionStore(), []);
  const syncService = useMemo(() => new SyncService(props.context, LIST_GUID), [props.context]);

  useEffect(() => {
    return NetworkService.subscribe((s) => {
      setNetwork(s);

      // When connectivity is stable again, attempt to sync queued offline submissions.
      if (s.online && !s.isWeak) {
        syncService.syncPending().catch((e) => {
          const msg = e instanceof Error ? e.message : String(e);
          Log.error(LOG_SOURCE, e instanceof Error ? e : new Error(msg));
        });
      }
    });
  }, [syncService]);

  useEffect(() => {
    const ctx = props.context as IContextWithItem;
    const row = ctx.item;
    if (!row) return;

    setIssue((prev) => ({
      ...prev,
      Id: row.Id ?? prev.Id,
      Title: row.Title ?? prev.Title,
      Description: row.Description ?? prev.Description,
      Priority: row.Priority ?? prev.Priority,
      Status: row.Status ?? prev.Status
    }));

    if (typeof row.RelatedIssueId === 'number') setRelatedIssueId(row.RelatedIssueId);

    let cancelled = false;
    if (row.Assignedto0Id) {
      services
        .getUserById(row.Assignedto0Id)
        .then((u) => {
          if (cancelled || !u) return;
          const text = u.Title || u.EMail || String(u.Id);
          setAssignedPeople([{ id: String(u.Id), text, secondaryText: u.EMail }]);
        })
        .catch(() => {
          // ignore prefill errors
        });
    }

    return () => {
      cancelled = true;
    };
  }, [props.context, services]);

  useEffect(() => {
    let cancelled = false;

    services
      .getItemTitles(RELATED_ISSUE_LIST_GUID, 200)
      .then((items) => {
        if (cancelled) return;
        const options: IDropdownOption[] = items.map((i) => ({ key: i.Id, text: i.Title || String(i.Id) }));
        setRelatedIssueOptions(options);
      })
      .catch(() => {
        // ignore load errors
      });

    return () => {
      cancelled = true;
    };
  }, [services]);

  const resolvePeopleSuggestions = useCallback(async (
    filterText: string,
    currentSelections?: IPersonaProps[]
  ): Promise<IPersonaProps[]> => {
    const users = await services.searchUsers(filterText, 10);
    const personas: IPersonaProps[] = users.map((u) => ({
      id: String(u.Id),
      text: u.Title || u.EMail || String(u.Id),
      secondaryText: u.EMail
    }));

    const selected = currentSelections || [];
    return personas.filter((p) => !selected.some((s) => s.id === p.id));
  }, [services]);

  const onSubmit = async (ev: React.FormEvent<HTMLFormElement>): Promise<void> => {
    ev.preventDefault();
    setError(undefined);
    setSuccess(undefined);
    setIsSaving(true);

    try {
      const selectedUserIdRaw = assignedPeople[0]?.id;
      const selectedUserId = selectedUserIdRaw ? Number(selectedUserIdRaw) : undefined;

      const payload: IIssueItem = {
        ...issue,
        Assignedto0: selectedUserId ? { Id: selectedUserId } : undefined,
        RelatedIssue: typeof relatedIssueId === 'number' ? { Id: relatedIssueId } : undefined
      };

      const networkState = NetworkService.getCurrentState();
      const shouldStoreOffline = !networkState.online || networkState.isWeak;

      const ctxWithItem = props.context as IContextWithItem;
      const itemIdForUpdate: number | undefined = payload.Id ?? ctxWithItem.itemId ?? ctxWithItem.item?.Id;

      if (props.displayMode === FormDisplayMode.Edit) {
        if (!itemIdForUpdate) {
          throw new Error('Missing item ID for update.');
        }

        if (shouldStoreOffline) {
          await offlineStore.enqueue({
            createdAt: Date.now(),
            webUrl: props.context.pageContext.web.absoluteUrl,
            listGuid: LIST_GUID,
            operation: 'update',
            itemId: itemIdForUpdate,
            payload
          });
          setSuccess('Saved locally (offline). It will sync when connectivity is restored.');
          return;
        }

        try {
          await services.updateItem(itemIdForUpdate, payload);
          setSuccess('Item has been saved.');
        } catch (e) {
          // Fallback: if a network error happens mid-save, keep the user’s work offline.
          const fallbackState = NetworkService.getCurrentState();
          if (!fallbackState.online || fallbackState.isWeak) {
            await offlineStore.enqueue({
              createdAt: Date.now(),
              webUrl: props.context.pageContext.web.absoluteUrl,
              listGuid: LIST_GUID,
              operation: 'update',
              itemId: itemIdForUpdate,
              payload
            });
            setSuccess('Saved locally (offline). It will sync when connectivity is restored.');
            return;
          }
          throw e;
        }
      } else if (props.displayMode === FormDisplayMode.New) {
        if (shouldStoreOffline) {
          await offlineStore.enqueue({
            createdAt: Date.now(),
            webUrl: props.context.pageContext.web.absoluteUrl,
            listGuid: LIST_GUID,
            operation: 'insert',
            payload
          });
          setSuccess('Saved locally (offline). It will sync when connectivity is restored.');
          return;
        }

        try {
          const created = await services.insertItem(payload);
          setIssue((prev) => ({ ...prev, Id: created.Id }));
          setSuccess('Item has been saved.');
        } catch (e) {
          const fallbackState = NetworkService.getCurrentState();
          if (!fallbackState.online || fallbackState.isWeak) {
            await offlineStore.enqueue({
              createdAt: Date.now(),
              webUrl: props.context.pageContext.web.absoluteUrl,
              listGuid: LIST_GUID,
              operation: 'insert',
              payload
            });
            setSuccess('Saved locally (offline). It will sync when connectivity is restored.');
            return;
          }
          throw e;
        }
      }
    } catch (e) {
      const message = e instanceof Error ? e.message : String(e);
      setError(message);
      const errorObj = e instanceof Error ? e : new Error(message);
      Log.error(LOG_SOURCE, errorObj);
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <>
      <div className={styles.offlineSpForm}>
        <form onSubmit={onSubmit}>
          <Stack
            tokens={{ childrenGap: 12 }}
            className={styles.mainStack}
          >
            <Stack
              horizontal
              verticalAlign="center"
              tokens={{ childrenGap: 8 }}
              className={styles.statusRow}
              style={{
                ['--offlineSpForm-statusDotColor' as unknown as keyof React.CSSProperties]:
                  isStableOnline ? theme.semanticColors.successText : theme.semanticColors.errorText
              }}
            >
              <button
                type="button"
                disabled
                className={styles.statusDotButton}
                aria-label={isStableOnline ? 'Online' : 'Offline'}
                title={isStableOnline ? 'Online' : 'Offline'}
              />
              <Label>{isStableOnline ? 'Online' : 'Offline'}</Label>
            </Stack>

            {success && (
              <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
                {success}
              </MessageBar>
            )}

            {error && (
              <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
                {error}
              </MessageBar>
            )}

            <Stack>
              <Label htmlFor="issueTitle">Title</Label>
              <TextField
                id="issueTitle"
                value={issue.Title}
                disabled={isReadOnly || isSaving}
                required
                onChange={(_, newValue) => setIssue((prev) => ({ ...prev, Title: newValue ?? '' }))}
              />
            </Stack>

            <Stack>
              <Label htmlFor="issueDescription">Description</Label>
              <TextField
                id="issueDescription"
                value={issue.Description ?? ''}
                disabled={isReadOnly || isSaving}
                multiline
                onChange={(_, newValue) => setIssue((prev) => ({ ...prev, Description: newValue ?? '' }))}
              />
            </Stack>

            <Stack>
              <Label htmlFor="issuePriority">Priority</Label>
              <Dropdown
                id="issuePriority"
                selectedKey={issue.Priority || undefined}
                disabled={isReadOnly || isSaving}
                options={PRIORITY_OPTIONS}
                onChange={(_, option) => setIssue((prev) => ({ ...prev, Priority: String(option?.key ?? '') }))}
              />
            </Stack>

            <Stack>
              <Label htmlFor="issueStatus">Status</Label>
              <Dropdown
                id="issueStatus"
                selectedKey={issue.Status || undefined}
                disabled={isReadOnly || isSaving}
                options={STATUS_OPTIONS}
                onChange={(_, option) => setIssue((prev) => ({ ...prev, Status: String(option?.key ?? '') }))}
              />
            </Stack>

            <Stack>
              <Label htmlFor="issueAssignedTo">Assigned To</Label>
              <NormalPeoplePicker
                inputProps={{ id: 'issueAssignedTo' }}
                selectedItems={assignedPeople}
                itemLimit={1}
                disabled={isReadOnly || isSaving}
                onResolveSuggestions={resolvePeopleSuggestions}
                getTextFromItem={(persona) => persona.text || ''}
                onChange={(items) => setAssignedPeople(items || [])}
                resolveDelay={250}
              />
            </Stack>

            <Stack>
              <Label htmlFor="issueRelated">Related Issue</Label>
              <Dropdown
                id="issueRelated"
                selectedKey={relatedIssueId}
                disabled={isReadOnly || isSaving}
                options={relatedIssueOptions}
                onChange={(_, option) => setRelatedIssueId(typeof option?.key === 'number' ? option.key : undefined)}
                placeholder="Select a related issue"
              />
            </Stack>

            <Stack horizontal tokens={{ childrenGap: 8 }}>
              {props.displayMode !== FormDisplayMode.Display && (
                <PrimaryButton type="submit" disabled={isSaving}>
                  {isSaving ? 'Saving...' : (isWeakNetwork ? 'Save (Offline)' : 'Save')}
                </PrimaryButton>
              )}

              <DefaultButton type="button" onClick={props.onClose} disabled={isSaving}>
                {isReadOnly ? 'Close' : 'Cancel'}
              </DefaultButton>
            </Stack>
          </Stack>
        </form>
      </div>
    </>
  );
}

export default OfflineSpForm;
