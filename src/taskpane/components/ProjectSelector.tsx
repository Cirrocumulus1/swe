import * as React from "react";
import { useEffect, useState } from "react";
import {
  Dropdown,
  Field,
  Option,
  Spinner,
  Text,
  makeStyles,
  tokens,
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
    rowGap: "10px",
    padding: "20px",
  },
  dropdown: {
    width: "100%",
  },
  selection: {
    color: tokens.colorNeutralForeground2,
  },
});

const ProjectSelector: React.FC<ProjectSelectorProps> = ({ onProjectSelected }) => {
  const styles = useStyles();
  const [projects, setProjects] = useState<Project[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [selectedProject, setSelectedProject] = useState<Project | undefined>(undefined);

  useEffect(() => {
    let isMounted = true;
    getProjects().then((result) => {
      if (isMounted) {
        setProjects(result);
        setIsLoading(false);
      }
    });
    return () => {
      isMounted = false;
    };
  }, []);

  const handleOptionSelect = (_event: SelectionEvents, data: OptionOnSelectData) => {
    const project = projects.find((item) => item.id === data.optionValue);
    setSelectedProject(project);
    if (project && onProjectSelected) {
      onProjectSelected(project);
    }
  };

  if (isLoading) {
    return <Spinner label="Loading projects..." />;
  }

  return (
    <div className={styles.container}>
      <Field label="Project">
        <Dropdown
          className={styles.dropdown}
          placeholder="Select a project"
          selectedOptions={selectedProject ? [selectedProject.id] : []}
          value={selectedProject ? selectedProject.name : ""}
          onOptionSelect={handleOptionSelect}
        >
          {projects.map((project) => (
            <Option key={project.id} value={project.id} text={project.name}>
              {`${project.name} (${project.customer})`}
            </Option>
          ))}
        </Dropdown>
      </Field>
      {selectedProject && (
        <Text className={styles.selection}>
          {`Selected: ${selectedProject.id} — ${selectedProject.name} · ${selectedProject.customer}`}
        </Text>
      )}
    </div>
  );
};

export default ProjectSelector;
