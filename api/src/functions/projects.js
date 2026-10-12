import { app } from "@azure/functions";

const CACHE_TTL_MS = 15 * 60 * 1000;
let cachedProjects;
let cacheExpiresAt = 0;

function jsonResponse(status, body) {
  return {
    status,
    headers: {
      "Cache-Control": "no-store",
      "Content-Type": "application/json; charset=utf-8",
    },
    jsonBody: body,
  };
}

async function fetchTeamspacePage(offset, username, password) {
  const baseUrl = process.env.TEAMSPACE_API_BASE_URL;
  if (!baseUrl || !username || !password) {
    throw new Error("Teamspace API settings are incomplete.");
  }

  const url = offset === 0
    ? `${baseUrl}?isActive=true&isTemplate=false`
    : `${baseUrl};offset=${offset}?isActive=true&isTemplate=false`;
  const authorization = Buffer.from(`${username}:${password}`).toString("base64");
  const response = await fetch(url, {
    headers: {
      Authorization: `Basic ${authorization}`,
      Accept: "application/json",
    },
  });

  if (!response.ok) {
    throw new Error(`Teamspace returned HTTP ${response.status}.`);
  }

  return response.json();
}

async function loadProjects() {
  const username = process.env.TEAMSPACE_USERNAME;
  const password = process.env.TEAMSPACE_PASSWORD;
  const firstPage = await fetchTeamspacePage(0, username, password);
  const allItems = [...(firstPage.items || [])];
  const totalCount = Number(firstPage.size) || 0;
  const pageSize = Number(firstPage.limit) || 100;

  if (pageSize < 1) {
    throw new Error("Teamspace returned an invalid page size.");
  }

  const offsets = [];
  for (let offset = pageSize; offset < totalCount; offset += pageSize) {
    offsets.push(offset);
  }

  const pages = await Promise.all(
    offsets.map((offset) => fetchTeamspacePage(offset, username, password))
  );
  for (const page of pages) {
    if (Array.isArray(page.items)) allItems.push(...page.items);
  }

  return allItems
    .filter((item) => item.path && !item.path.includes("/"))
    .map((item) => ({
      id: String(item.value),
      name: item.caption,
      customer: item.caption,
      path: item.path,
    }))
    .sort((left, right) => left.name.localeCompare(right.name, "de", { sensitivity: "base" }));
}

async function getProjects() {
  if (cachedProjects && Date.now() < cacheExpiresAt) {
    return cachedProjects;
  }

  const staleProjects = cachedProjects;
  try {
    cachedProjects = await loadProjects();
    cacheExpiresAt = Date.now() + CACHE_TTL_MS;
    return cachedProjects;
  } catch (error) {
    if (staleProjects) return staleProjects;
    throw error;
  }
}

app.http("projects", {
  methods: ["GET"],
  authLevel: "anonymous",
  route: "projects",
  handler: async (_request, context) => {
    try {
      return jsonResponse(200, await getProjects());
    } catch (error) {
      context.error("Project API could not load the Teamspace project list.", error);
      return jsonResponse(502, { error: "The project list is temporarily unavailable." });
    }
  },
});
