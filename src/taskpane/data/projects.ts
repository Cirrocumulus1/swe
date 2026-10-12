/* global window */

export interface Project {
  id: string;
  name: string;
  customer: string;
  path: string;
}

interface CachedProjects {
  timestamp: number;
  data: Project[];
}

const CACHE_KEY = "teamspace_projects_cache";
const CACHE_TTL_MS = 4 * 60 * 60 * 1000; // 4 Stunden Gültigkeit
const PROJECTS_API_URL = "/api/projects";

class ProjectApiError extends Error {
  constructor(
    message: string,
    readonly status: number
  ) {
    super(message);
    this.name = "ProjectApiError";
  }
}

function readCachedProjects(): CachedProjects | undefined {
  try {
    const cachedRaw = window.localStorage.getItem(CACHE_KEY);
    if (!cachedRaw) return undefined;

    const cached: CachedProjects = JSON.parse(cachedRaw);
    if (!Number.isFinite(cached.timestamp) || !Array.isArray(cached.data)) return undefined;
    return cached;
  } catch {
    return undefined;
  }
}

async function fetchFromApi(): Promise<Project[]> {
  const response = await window.fetch(PROJECTS_API_URL, {
    method: "GET",
    headers: { Accept: "application/json" },
  });

  if (!response.ok) {
    throw new ProjectApiError(
      `Projektliste konnte nicht geladen werden (HTTP ${response.status}).`,
      response.status
    );
  }

  return response.json();
}

export async function getProjects(forceRefresh = false): Promise<Project[]> {
  if (!forceRefresh) {
    const cached = readCachedProjects();
    if (cached && Date.now() - cached.timestamp < CACHE_TTL_MS) {
      return cached.data;
    }
  }

  try {
    const freshProjects = await fetchFromApi();
    try {
      const toCache: CachedProjects = {
        timestamp: Date.now(),
        data: freshProjects,
      };
      window.localStorage.setItem(CACHE_KEY, JSON.stringify(toCache));
    } catch {
      // Continue without a persistent cache if storage is unavailable.
    }
    return freshProjects;
  } catch (error) {
    const staleCache = readCachedProjects();
    if (staleCache) {
      return staleCache.data;
    }

    throw error;
  }
}
