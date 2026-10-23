import * as React from "react";
import { useEffect, useState } from "react";
import Header from "./Header";
import ProjectSelector from "./ProjectSelector";
import {
  Badge,
  Button,
  Card,
  MessageBar,
  MessageBarBody,
  Spinner,
  Text,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { ArrowSyncRegular } from "@fluentui/react-icons";
import { Project } from "../data/projects";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    display: "flex",
    flexDirection: "column",
    backgroundColor: tokens.colorNeutralBackground1,
  },
  existingTagCard: {
    margin: "12px 20px 0 20px",
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground2,
    display: "flex",
    flexDirection: "column",
    rowGap: "6px",
  },
  statusContainer: {
    padding: "12px 20px 20px 20px",
    display: "flex",
    flexDirection: "column",
    rowGap: "8px",
  },
  actionContainer: {
    padding: "0 20px",
  },
  syncResults: {
    maxHeight: "180px",
    overflowY: "auto",
    margin: "8px 0 0",
    paddingLeft: "20px",
  },
});

interface ExistingTag {
  id: string;
  name?: string;
  categoryTag: string;
  assignedAt?: string;
}

interface MessageSelection {
  itemIds: string[];
  subjects: string[];
}

interface SyncItemResult {
  subject: string;
  status: "succeeded" | "failed";
  error?: string;
}

function formatAssignmentTimestamp(value?: string): string | undefined {
  if (!value) return undefined;

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return undefined;

  return date.toLocaleString("de-DE");
}

function saveProjectProperties(
  item: NonNullable<typeof Office.context.mailbox.item>,
  project: Project,
  assignedAt: string
): Promise<void> {
  return new Promise((resolve, reject) => {
    item.loadCustomPropertiesAsync((propResult) => {
      if (propResult.status === Office.AsyncResultStatus.Failed) {
        reject(new Error(`Fehler beim Laden der Metadaten: ${propResult.error.message}`));
        return;
      }

      const customProps = propResult.value;
      customProps.set("teamspace_project_id", project.id);
      customProps.set("teamspace_project_name", project.name);
      customProps.set("teamspace_tagged_at", assignedAt);
      customProps.saveAsync((saveResult) => {
        if (saveResult.status === Office.AsyncResultStatus.Failed) {
          reject(new Error(`Fehler beim Speichern der Projekt-Metadaten: ${saveResult.error.message}`));
          return;
        }
        resolve();
      });
    });
  });
}

function getMessageItemId(item: Office.Item): Promise<string> {
  const composeItem = item as Office.MessageCompose;
  if (typeof composeItem.saveAsync === "function") {
    return new Promise((resolve, reject) => {
      composeItem.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(new Error(`Entwurf konnte nicht gespeichert werden: ${result.error.message}`));
        } else if (result.value) {
          resolve(result.value);
        } else {
          reject(new Error("Outlook hat keine Nachrichten-ID zurückgegeben."));
        }
      });
    });
  }

  const itemId = (item as Office.MessageRead).itemId;
  if (!itemId) return Promise.reject(new Error("Die Nachrichten-ID ist nicht verfügbar."));
  return Promise.resolve(itemId);
}

function getSelectedMessageItems(): Promise<Office.SelectedItemDetails[]> {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.getSelectedItemsAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        reject(new Error(`Ausgewählte Nachrichten konnten nicht abgerufen werden: ${result.error.message}`));
        return;
      }
      resolve(result.value || []);
    });
  });
}

function convertToRestMessageId(itemId: string): string {
  try {
    return Office.context.mailbox.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0);
  } catch (error) {
    console.error("Die Outlook-Nachrichten-ID konnte nicht ins REST-Format konvertiert werden.", error);
    throw error;
  }
}

async function getRestMessageSelection(item?: Office.Item): Promise<MessageSelection> {
  const isCompose = item && typeof (item as Office.MessageCompose).saveAsync === "function";
  const mailbox = Office.context.mailbox;

  if (!isCompose && typeof mailbox.getSelectedItemsAsync === "function") {
    let selectedItems: Office.SelectedItemDetails[];
    try {
      selectedItems = await getSelectedMessageItems();
    } catch (error) {
      console.error("Die ausgewählten Outlook-Nachrichten konnten nicht abgerufen werden.", error);
      throw error;
    }

    if (selectedItems.length > 1 || (!item && selectedItems.length === 1)) {
      return {
        itemIds: selectedItems.map((selectedItem) => convertToRestMessageId(selectedItem.itemId)),
        subjects: selectedItems.map((selectedItem) => selectedItem.subject || ""),
      };
    }
  }

  if (!item) throw new Error("Es wurde keine Outlook-Nachricht ausgewählt.");

  let currentItemId: string;
  try {
    currentItemId = await getMessageItemId(item);
  } catch (error) {
    console.error("Die Outlook-Nachrichten-ID konnte nicht abgerufen werden.", error);
    throw error;
  }

  return {
    itemIds: [convertToRestMessageId(currentItemId)],
    subjects: [await getMessageSubject(item)],
  };
}

function getMessageSubject(item: Office.Item): Promise<string> {
  const composeSubject = (item as Office.MessageCompose).subject as unknown;
  if (composeSubject && typeof composeSubject === "object" && "getAsync" in composeSubject) {
    return new Promise((resolve) => {
      (composeSubject as Office.Subject).getAsync((result) => {
        resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : "");
      });
    });
  }

  const readSubject = (item as Office.MessageRead).subject;
  return Promise.resolve(typeof readSubject === "string" ? readSubject : "");
}

const App: React.FC<{ title: string }> = ({ title }) => {
  const styles = useStyles();
  const [existingTag, setExistingTag] = useState<ExistingTag | null>(null);
  const [statusMessage, setStatusMessage] = useState<{
    text: string;
    intent: "success" | "error";
    items?: SyncItemResult[];
  } | null>(null);
  const [selectedProject, setSelectedProject] = useState<Project | null>(null);
  const [isSyncing, setIsSyncing] = useState<boolean>(false);

  // Bestehende Zuweisungen beim Öffnen der Mail prüfen (Primär über customProperties)
  useEffect(() => {
    if (!window.Office?.context?.mailbox?.item) return;

    const item = Office.context.mailbox.item;

    // 1. Zuerst aus customProperties lesen (kompatibel mit Neuem Outlook / Web im Compose-Modus)
    item.loadCustomPropertiesAsync((propResult) => {
      if (propResult.status === Office.AsyncResultStatus.Succeeded) {
        const props = propResult.value;
        const projId = props.get("teamspace_project_id");
        const projName = props.get("teamspace_project_name");
        const assignedAt = props.get("teamspace_tagged_at");

        if (projId) {
          setExistingTag({
            id: String(projId),
            name: projName ? String(projName) : undefined,
            categoryTag: `TS-${projId}`,
            assignedAt: typeof assignedAt === "string" ? assignedAt : undefined,
          });
          return;
        }
      }

      /* 
       * HINWEIS ZUR EXCHANGE-KATEGORIEN-API:
       * Im Neuen Outlook und OWA blockiert Microsoft item.categories im Compose-Modus.
       * 
       * if (item.categories) {
       *   item.categories.getAsync((catResult) => {
       *     if (catResult.status === Office.AsyncResultStatus.Succeeded) {
       *       const categories = catResult.value || [];
       *       const tsCategory = categories.find((c: any) => {
       *         const name = typeof c === "string" ? c : c.displayName;
       *         return name && name.startsWith("TS-");
       *       });
       *       if (tsCategory) {
       *         const rawName = typeof tsCategory === "string" ? tsCategory : tsCategory.displayName;
       *         setExistingTag({
       *           id: rawName.replace("TS-", ""),
       *           categoryTag: rawName,
       *         });
       *       }
       *     }
       *   });
       * }
       */
    });
  }, []);

  /*
   * =========================================================================================
   * INAKTIV: Kategorien-Zuweisung via Exchange masterCategories & item.categories
   * Microsoft wirft im Neuen Outlook / OWA im Compose-Modus: "Dieser Vorgang wird nicht unterstützt."
   * =========================================================================================
   *
   * const ensureCategoryAndAssign = (categoryTag: string, onComplete: (error?: string) => void) => {
   *   const mailbox = Office.context.mailbox;
   *   const item = mailbox.item;
   *   if (!mailbox.masterCategories) {
   *     item.categories.addAsync([categoryTag], (res) => {
   *       if (res.status === Office.AsyncResultStatus.Failed) onComplete(res.error.message);
   *       else onComplete();
   *     });
   *     return;
   *   }
   *   mailbox.masterCategories.getAsync((masterResult) => {
   *     if (masterResult.status === Office.AsyncResultStatus.Failed) {
   *       onComplete(`Hauptkategorien nicht lesbar: ${masterResult.error.message}`);
   *       return;
   *     }
   *     const masterList = masterResult.value || [];
   *     const existsInMaster = masterList.some((cat: any) => cat.displayName === categoryTag);
   *     const assignToItem = () => {
   *       item.categories.addAsync([categoryTag], (assignResult) => {
   *         if (assignResult.status === Office.AsyncResultStatus.Failed) onComplete(assignResult.error.message);
   *         else onComplete();
   *       });
   *     };
   *     if (!existsInMaster) {
   *       const newCategory = {
   *         displayName: categoryTag,
   *         color: (Office.MailboxEnums?.CategoryColor?.Preset0 ?? "Preset0") as Office.MailboxEnums.CategoryColor,
   *       };
   *       mailbox.masterCategories.addAsync([newCategory], (addMasterResult) => {
   *         if (addMasterResult.status === Office.AsyncResultStatus.Failed) {
   *           onComplete(`Kategorie konnte nicht in Hauptliste angelegt werden: ${addMasterResult.error.message}`);
   *           return;
   *         }
   *         assignToItem();
   *       });
   *     } else {
   *       assignToItem();
   *     }
   *   });
   * };
   */

  const handleProjectSelected = (project: Project) => {
    setSelectedProject(project);
    setStatusMessage(null);
  };

  const handleSyncSelected = async () => {
    if (!selectedProject) return;
    if (!window.Office?.context?.mailbox) {
      setStatusMessage({ text: "Office Mailbox Kontext nicht verfügbar.", intent: "error" });
      return;
    }

    setIsSyncing(true);
    setStatusMessage(null);

    try {
      const mailbox = Office.context.mailbox;
      const item = mailbox.item || undefined;
      const selection = await getRestMessageSelection(item);
      const project = selectedProject;
      if (selection.itemIds.length === 1 && item) {
        const assignedAt = new Date().toISOString();
        await saveProjectProperties(item, project, assignedAt);
        setExistingTag({
          id: project.id,
          name: project.name,
          categoryTag: `TS-${project.id}`,
          assignedAt,
        });
      } else {
        setExistingTag(null);
      }

      const userEmail = mailbox.userProfile.emailAddress;
      const accessToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true });
      const response = await window.fetch("/api/sync-email", {
        method: "POST",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
          Accept: "application/json",
        },
        body: JSON.stringify(
          selection.itemIds.length === 1
            ? {
                userEmail,
                itemId: selection.itemIds[0],
                teamspaceId: project.id,
                subject: selection.subjects[0] || "",
              }
            : {
                userEmail,
                itemIds: selection.itemIds,
                teamspaceId: project.id,
                subjects: selection.subjects,
              }
        ),
      });

      const result = (await response.json().catch(() => ({}))) as {
        message?: string;
        error?: string;
        items?: SyncItemResult[];
      };
      if (!response.ok) {
        setStatusMessage({
          text: result.error || `Synchronisierung fehlgeschlagen (HTTP ${response.status}).`,
          intent: "error",
          items: result.items,
        });
        return;
      }

      const items = Array.isArray(result.items) ? result.items : [];
      setStatusMessage({
        text: result.message || `E-Mail erfolgreich mit "${project.name}" synchronisiert.`,
        intent: items.some((syncItem) => syncItem.status === "failed") ? "error" : "success",
        items,
      });
    } catch (error) {
      setStatusMessage({
        text: error instanceof Error ? error.message : "E-Mail konnte nicht synchronisiert werden.",
        intent: "error",
      });
    } finally {
      setIsSyncing(false);
    }
  };

  return (
    <div className={styles.root}>
      <Header logo="assets/logo-filled.png" title={title} />

      {existingTag && (
        <Card className={styles.existingTagCard}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <Text weight="semibold">Zugeordnetes Projekt:</Text>
            <Badge appearance="filled" color="brand">
              {existingTag.categoryTag}
            </Badge>
          </div>
          {existingTag.name && <Text size={300}>{existingTag.name}</Text>}
          <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
            Teamspace-ID: {existingTag.id}
          </Text>
          {formatAssignmentTimestamp(existingTag.assignedAt) && (
            <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
              Zugeordnet am: {formatAssignmentTimestamp(existingTag.assignedAt)}
            </Text>
          )}
        </Card>
      )}

      <ProjectSelector onProjectSelected={handleProjectSelected} />

      <div className={styles.actionContainer}>
        <Button
          appearance="primary"
          icon={<ArrowSyncRegular />}
          disabled={!selectedProject || isSyncing}
          onClick={handleSyncSelected}
        >
          Markieren &amp; Sync
        </Button>
      </div>

      <div className={styles.statusContainer}>
        {isSyncing && <Spinner size="tiny" label="Synchronisiere E-Mail..." />}

        {statusMessage && !isSyncing && (
          <MessageBar intent={statusMessage.intent}>
            <MessageBarBody>
              {statusMessage.text}
              {statusMessage.items && statusMessage.items.length > 0 && (
                <ul className={styles.syncResults}>
                  {statusMessage.items.map((syncItem, index) => (
                    <li key={`${syncItem.subject}-${index}`}>
                      {syncItem.subject || `E-Mail ${index + 1}`}: {" "}
                      {syncItem.status === "succeeded"
                        ? "synchronisiert"
                        : `fehlgeschlagen${syncItem.error ? `: ${syncItem.error}` : ""}`}
                    </li>
                  ))}
                </ul>
              )}
            </MessageBarBody>
          </MessageBar>
        )}
      </div>
    </div>
  );
};

export default App;
