import * as React from "react";
import { useEffect, useState, useMemo } from "react";
import {
  Button,
  Combobox,
  Field,
  MessageBar,
  MessageBarBody,
  Option,
  Spinner,
  makeStyles,
  type OptionOnSelectData,
  type SelectionEvents,
} from "@fluentui/react-components";
import { getProjects, Project } from "../data/projects";

interface ProjectSelectorProps {
  onProjectSelected?: (project: Project) => void;
}

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    rowGap: "12px",
    padding: "20px",
  },
  labelRow: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    width: "100%",
  },
  combobox: {
    width: "100%",
  },
  listbox: {
    maxHeight: "min(240px, 40vh)",
    overflowY: "auto",
    overscrollBehaviorY: "contain",
  },
});

// Schlankes Inline-SVG – spart externe Icon-Pakete
const RefreshIcon: React.FC = () => (
  <svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor">
    <path d="M13.65 2.35A8 8 0 1 0 16 8h-1.5a6.5 6.5 0 1 1-1.9-4.6L10 6h6V0l-2.35 2.35z" />
  </svg>
);

const ProjectSelector: React.FC<ProjectSelectorProps> = ({ onProjectSelected }) => {
  const styles = useStyles();
  const [projects, setProjects] = useState<Project[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [searchTerm, setSearchTerm] = useState<string>("");
  const [selectedProject, setSelectedProject] = useState<Project | undefined>(undefined);

  const loadProjectData = (force = false) => {
    setIsLoading(true);
    setLoadError(null);
    getProjects(force)
      .then((result) => {
        setProjects(result);
      })
      .catch((error: unknown) => {
        setLoadError(error instanceof Error ? error.message : "Projektliste konnte nicht geladen werden.");
      })
      .finally(() => {
        setIsLoading(false);
      });
  };

  useEffect(() => {
    loadProjectData(false);
  }, []);

  const filteredProjects = useMemo(() => {
    if (!searchTerm.trim()) return projects;
    const query = searchTerm.toLowerCase();
    return projects.filter(
      (p) => p.name.toLowerCase().includes(query) || p.id.toLowerCase().includes(query)
    );
  }, [projects, searchTerm]);

  const handleOptionSelect = (_event: SelectionEvents, data: OptionOnSelectData) => {
    const project = projects.find((item) => item.id === data.optionValue);
    setSelectedProject(project);
    setSearchTerm(project ? project.name : "");
    if (project && onProjectSelected) {
      onProjectSelected(project);
    }
  };

  const handleInputChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    setSearchTerm(event.target.value);
  };

  if (isLoading && projects.length === 0) {
    return <Spinner label="Lade Teamspace-Projekte..." />;
  }

  return (
    <div className={styles.container}>
      {loadError && (
        <MessageBar intent="error">
          <MessageBarBody>{loadError}</MessageBarBody>
        </MessageBar>
      )}
      <Field
        label={
          <div className={styles.labelRow}>
            <span>Teamspace Projekt suchen</span>
            <Button
              appearance="subtle"
              size="small"
              icon={<RefreshIcon />}
              title="Projektliste neu laden"
              disabled={isLoading}
              onClick={() => loadProjectData(true)}
            />
          </div>
        }
      >
        <Combobox
          className={styles.combobox}
          placeholder="Projektname eingeben..."
          value={searchTerm}
          onChange={handleInputChange}
          onOptionSelect={handleOptionSelect}
          selectedOptions={selectedProject ? [selectedProject.id] : []}
          positioning={{ position: "below", autoSize: false }}
          listbox={{
            className: styles.listbox,
            style: {
              maxHeight: "min(240px, 40vh)",
              overflowY: "auto",
              overscrollBehaviorY: "contain",
            },
          }}
        >
          {filteredProjects.length === 0 ? (
            <Option disabled value="no-results">
              Kein Projekt gefunden
            </Option>
          ) : (
            filteredProjects.map((project) => (
              <Option key={project.id} value={project.id} text={project.name}>
                {project.name}
              </Option>
            ))
          )}
        </Combobox>
      </Field>
    </div>
  );
};

export default ProjectSelector;