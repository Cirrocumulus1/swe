export interface Project {
  id: string;
  name: string;
  customer: string;
}

// Stubbed data until the project API is available.
export const projects: Project[] = [
  { id: "P-1001", name: "Website Relaunch", customer: "Contoso" },
  { id: "P-1002", name: "ERP Migration", customer: "Fabrikam" },
  { id: "P-1003", name: "Data Warehouse", customer: "Northwind" },
  { id: "P-1004", name: "Mobile App", customer: "Adventure Works" },
  { id: "P-1005", name: "Security Audit", customer: "Tailspin Toys" },
];

export async function getProjects(): Promise<Project[]> {
  return projects;
}
