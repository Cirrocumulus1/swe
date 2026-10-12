import * as React from "react";
import { useEffect, useState } from "react";
import Header from "./Header";
import ProjectSelector from "./ProjectSelector";
import {
  Badge,
  Card,
  MessageBar,
  MessageBarBody,
  Spinner,
  Text,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
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
});

interface ExistingTag {
  id: string;
  name?: string;
  categoryTag: string;
}

const App: React.FC<{ title: string }> = ({ title }) => {
  const styles = useStyles();
  const [existingTag, setExistingTag] = useState<ExistingTag | null>(null);
  const [statusMessage, setStatusMessage] = useState<{
    text: string;
    intent: "success" | "error";
  } | null>(null);
  const [isSaving, setIsSaving] = useState<boolean>(false);

  // Bestehende Zuweisungen beim Öffnen der Mail prüfen
  useEffect(() => {
    if (window.Office?.context?.mailbox?.item?.categories) {
      const item = Office.context.mailbox.item;

      item.categories.getAsync((catResult) => {
        if (catResult.status === Office.AsyncResultStatus.Succeeded) {
          const categories = catResult.value || [];
          const tsCategory = categories.find((c: any) => {
            const name = typeof c === "string" ? c : c.displayName;
            return name && name.startsWith("TS-");
          });

          if (tsCategory) {
            const rawName = typeof tsCategory === "string" ? tsCategory : tsCategory.displayName;
            const extractedId = rawName.replace("TS-", "");

            item.loadCustomPropertiesAsync((propResult) => {
              const projName =
                propResult.status === Office.AsyncResultStatus.Succeeded
                  ? propResult.value.get("teamspace_project_name")
                  : undefined;

              setExistingTag({
                id: extractedId,
                name: projName ? String(projName) : undefined,
                categoryTag: rawName,
              });
            });
          }
        }
      });
    }
  }, []);

  // Stellt sicher, dass das Tag in der Hauptliste existiert, bevor es der Mail zugewiesen wird
  const ensureCategoryAndAssign = (categoryTag: string, onComplete: (error?: string) => void) => {
    const mailbox = Office.context.mailbox;
    const item = mailbox.item;

    if (!mailbox.masterCategories) {
      // Fallback für Umgebungen ohne masterCategories-Zugriff
      item.categories.addAsync([categoryTag], (res) => {
        if (res.status === Office.AsyncResultStatus.Failed) {
          onComplete(res.error.message);
        } else {
          onComplete();
        }
      });
      return;
    }

    // 1. Hauptkategorieliste abfragen
    mailbox.masterCategories.getAsync((masterResult) => {
      if (masterResult.status === Office.AsyncResultStatus.Failed) {
        onComplete(`Hauptkategorien nicht lesbar: ${masterResult.error.message}`);
        return;
      }

      const masterList = masterResult.value || [];
      const existsInMaster = masterList.some((cat: any) => cat.displayName === categoryTag);

      const assignToItem = () => {
        item.categories.addAsync([categoryTag], (assignResult) => {
          if (assignResult.status === Office.AsyncResultStatus.Failed) {
            onComplete(assignResult.error.message);
          } else {
            onComplete();
          }
        });
      };

      // 2. Falls noch nicht in der Hauptliste: Anlegen und danach zuweisen
      if (!existsInMaster) {
        const newCategory = {
          displayName: categoryTag,
          color: (Office.MailboxEnums?.CategoryColor?.Preset0 ?? "Preset0") as Office.MailboxEnums.CategoryColor,
        };

        mailbox.masterCategories.addAsync([newCategory], (addMasterResult) => {
          if (addMasterResult.status === Office.AsyncResultStatus.Failed) {
            onComplete(`Kategorie konnte nicht in Hauptliste angelegt werden: ${addMasterResult.error.message}`);
            return;
          }
          assignToItem();
        });
      } else {
        assignToItem();
      }
    });
  };

  const handleProjectSelected = (project: Project) => {
    if (!window.Office?.context?.mailbox?.item) {
      setStatusMessage({
        text: "Office Mailbox Kontext nicht verfügbar.",
        intent: "error",
      });
      return;
    }

    setIsSaving(true);
    setStatusMessage(null);

    const item = Office.context.mailbox.item;
    const categoryTag = `TS-${project.id}`;

    ensureCategoryAndAssign(categoryTag, (error) => {
      if (error) {
        setStatusMessage({
          text: `Fehler beim Setzen der Kategorie: ${error}`,
          intent: "error",
        });
        setIsSaving(false);
        return;
      }

      // Metadaten parallel persistieren
      item.loadCustomPropertiesAsync((propResult) => {
        if (propResult.status === Office.AsyncResultStatus.Succeeded) {
          const customProps = propResult.value;
          customProps.set("teamspace_project_id", project.id);
          customProps.set("teamspace_project_name", project.name);
          customProps.set("teamspace_tagged_at", new Date().toISOString());
          customProps.saveAsync();
        }

        setExistingTag({
          id: project.id,
          name: project.name,
          categoryTag: categoryTag,
        });

        setStatusMessage({
          text: `Kategorie "${categoryTag}" (${project.name}) zugewiesen.`,
          intent: "success",
        });
        setIsSaving(false);
      });
    });
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
        </Card>
      )}

      <ProjectSelector onProjectSelected={handleProjectSelected} />

      <div className={styles.statusContainer}>
        {isSaving && <Spinner size="tiny" label="Prüfe und weise Kategorie zu..." />}

        {statusMessage && !isSaving && (
          <MessageBar intent={statusMessage.intent}>
            <MessageBarBody>{statusMessage.text}</MessageBarBody>
          </MessageBar>
        )}
      </div>
    </div>
  );
};

export default App;