import { app } from "@azure/functions";
import { ClientSecretCredential } from "@azure/identity";
import { createRemoteJWKSet, jwtVerify } from "jose";

const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
const GRAPH_SCOPE = "https://graph.microsoft.com/.default";
let graphCredential;
const signingKeysByTenant = new Map();

class RequestError extends Error {
  constructor(status, message) {
    super(message);
    this.status = status;
  }
}

class GraphApiError extends Error {
  constructor(status, operation) {
    super(`Microsoft Graph ${operation} failed with HTTP ${status}.`);
    this.status = status;
    this.operation = operation;
  }
}

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

function requiredSetting(name) {
  const value = process.env[name]?.trim();
  if (!value) throw new Error(`The ${name} app setting is required.`);
  return value;
}

function getSigningKeys(tenantId) {
  let signingKeys = signingKeysByTenant.get(tenantId);
  if (!signingKeys) {
    const jwksUri = new URL(
      `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/discovery/v2.0/keys`
    );
    signingKeys = createRemoteJWKSet(jwksUri);
    signingKeysByTenant.set(tenantId, signingKeys);
  }
  return signingKeys;
}

async function validateOfficeToken(request, context) {
  const authorization = request.headers.get("authorization") || "";
  const match = authorization.match(/^Bearer\s+(.+)$/i);
  if (!match) throw new RequestError(401, "Anmeldung erforderlich. Bitte Outlook erneut öffnen.");

  const tenantId = requiredSetting("TENANT_ID");
  const clientId = requiredSetting("CLIENT_ID");
  const appIdUri = process.env.ADDIN_APP_ID_URI?.trim() ||
    `api://polite-sky-0bea56210.4.azurestaticapps.net/${clientId}`;
  const acceptedIssuers = [
    `https://login.microsoftonline.com/${tenantId}/v2.0`,
    `https://sts.windows.net/${tenantId}/`,
  ];
  const acceptedAudiences = [clientId, appIdUri, `api://${clientId}`];

  let tokenPayload;
  try {
    ({ payload: tokenPayload } = await jwtVerify(match[1].trim(), getSigningKeys(tenantId), {
      issuer: acceptedIssuers,
      audience: acceptedAudiences,
    }));
  } catch (error) {
    context.error("Outlook token verification failed:", error);
    throw new RequestError(401, "Die Outlook-Anmeldung konnte nicht geprüft werden. Bitte Outlook erneut öffnen.");
  }

  const requiredScope = process.env.ADDIN_SSO_SCOPE?.trim() || "access_as_user";
  const scopes = typeof tokenPayload.scp === "string" ? tokenPayload.scp.split(" ") : [];
  if (!scopes.includes(requiredScope)) {
    throw new RequestError(403, "Die Outlook-Anmeldung besitzt nicht die erforderliche Berechtigung.");
  }

  const callerEmail = [tokenPayload.preferred_username, tokenPayload.upn, tokenPayload.email]
    .find((value) => typeof value === "string" && value.includes("@"));
  if (!callerEmail) {
    throw new RequestError(401, "Die Outlook-Anmeldung enthält keine Benutzer-Mailadresse.");
  }

  return callerEmail.trim().toLowerCase();
}

function parsePayload(body) {
  if (!body || typeof body !== "object" || Array.isArray(body)) {
    throw new RequestError(400, "Ungültiger JSON-Request.");
  }

  const requiredFields = ["userEmail", "teamspaceId"];
  for (const field of requiredFields) {
    if (typeof body[field] !== "string" || !body[field].trim()) {
      throw new RequestError(400, `Pflichtfeld fehlt oder ist ungültig: ${field}.`);
    }
  }

  let itemIds;
  if (body.itemIds !== undefined) {
    if (!Array.isArray(body.itemIds) || body.itemIds.length === 0 || body.itemIds.length > 100) {
      throw new RequestError(400, "itemIds muss zwischen 1 und 100 Nachrichten-IDs enthalten.");
    }
    itemIds = body.itemIds.map((itemId) => {
      if (typeof itemId !== "string" || !itemId.trim() || itemId.length > 4096) {
        throw new RequestError(400, "Eine Nachrichten-ID ist ungültig.");
      }
      return itemId.trim();
    });
  } else if (typeof body.itemId === "string" && body.itemId.trim() && body.itemId.length <= 4096) {
    itemIds = [body.itemId.trim()];
  } else {
    throw new RequestError(400, "Pflichtfeld fehlt oder ist ungültig: itemId oder itemIds.");
  }

  if (body.userEmail.length > 320 || body.teamspaceId.length > 256) {
    throw new RequestError(400, "Ein Request-Feld überschreitet die zulässige Länge.");
  }

  const subject = typeof body.subject === "string" ? body.subject : "";
  if (subject.length > 1000) throw new RequestError(400, "Der Betreff überschreitet die zulässige Länge.");
  let subjects = itemIds.map(() => subject);
  if (body.subjects !== undefined) {
    if (
      !Array.isArray(body.subjects) ||
      body.subjects.length !== itemIds.length ||
      body.subjects.some((value) => typeof value !== "string" || value.length > 1000)
    ) {
      throw new RequestError(400, "subjects muss pro Nachrichten-ID einen gültigen Betreff enthalten.");
    }
    subjects = body.subjects;
  }

  return {
    userEmail: body.userEmail.trim(),
    itemIds,
    teamspaceId: body.teamspaceId.trim(),
    subjects,
  };
}

function getGraphCredential() {
  if (!graphCredential) {
    graphCredential = new ClientSecretCredential(
      requiredSetting("TENANT_ID"),
      requiredSetting("CLIENT_ID"),
      requiredSetting("CLIENT_SECRET")
    );
  }
  return graphCredential;
}

async function getGraphAccessToken() {
  const accessToken = await getGraphCredential().getToken(GRAPH_SCOPE);
  if (!accessToken?.token) throw new Error("Microsoft Graph access token was not issued.");
  return accessToken.token;
}

async function graphRequest(url, accessToken, operation, options = {}) {
  let response;
  try {
    response = await fetch(url, {
      ...options,
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: "application/json",
        ...options.headers,
      },
    });
  } catch {
    throw new GraphApiError(0, operation);
  }

  if (!response.ok) throw new GraphApiError(response.status, operation);
  return response;
}

function descriptionContainsTeamspaceId(description, teamspaceId) {
  if (typeof description !== "string") return false;
  const escapedId = teamspaceId.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const idPattern = new RegExp(`(^|[^A-Za-z0-9])${escapedId}($|[^A-Za-z0-9])`, "i");
  return idPattern.test(description);
}

async function resolveSharePointSite(teamspaceId, accessToken) {
  const searchUrl = new URL(`${GRAPH_BASE_URL}/sites`);
  searchUrl.searchParams.set("search", teamspaceId);
  searchUrl.searchParams.set("$select", "id,name,displayName,webUrl,description");

  const matchingSites = new Map();
  let pageUrl = searchUrl.toString();
  while (pageUrl) {
    const response = await graphRequest(pageUrl, accessToken, "SharePoint-Site-Suche");
    const result = await response.json();
    const sites = Array.isArray(result.value) ? result.value : [];
    for (const site of sites) {
      if (site.id && descriptionContainsTeamspaceId(site.description, teamspaceId)) {
        matchingSites.set(site.id, site);
      }
    }

    const nextLink = result["@odata.nextLink"];
    if (nextLink) {
      const nextUrl = new URL(nextLink);
      if (nextUrl.origin !== new URL(GRAPH_BASE_URL).origin) {
        throw new Error("Microsoft Graph returned an unexpected site-search URL.");
      }
      pageUrl = nextUrl.toString();
    } else {
      pageUrl = undefined;
    }
  }

  if (matchingSites.size === 0) {
    throw new RequestError(404, "Die SharePoint-Projektseite wurde nicht gefunden.");
  }
  if (matchingSites.size > 1) {
    throw new RequestError(409, "Die Teamspace-ID entspricht mehreren SharePoint-Seiten.");
  }
  return matchingSites.values().next().value;
}

async function verifyMailFolder(siteId, accessToken) {
  const folderUrl = new URL(
    `${GRAPH_BASE_URL}/sites/${encodeURIComponent(siteId)}/drive/root:/E-Mails`
  );
  folderUrl.searchParams.set("$select", "id,name,folder");

  let response;
  try {
    response = await graphRequest(folderUrl, accessToken, "SharePoint-Zielordner");
  } catch (error) {
    if (error instanceof GraphApiError && error.status === 404) {
      throw new RequestError(
        404,
        "Der SharePoint-Zielordner /E-Mails fehlt auf der gefundenen Projekt-Site. Lege den Ordner an und versuche es erneut."
      );
    }
    throw error;
  }

  const folder = await response.json();
  if (!folder.folder) {
    throw new RequestError(409, "Der SharePoint-Pfad /E-Mails ist vorhanden, aber kein Ordner.");
  }
}

function createFileName(subject, suffix = "") {
  const timestamp = new Date().toISOString().slice(0, 16).replace("T", "_").replace(":", "");
  const cleanSubject = (subject || "Ohne Betreff")
    .replace(/[\\/:*?"<>|#%&]/g, "-")
    .replace(/[\u0000-\u001f]/g, "-")
    .trim()
    .replace(/[. ]+$/g, "")
    .slice(0, 220) || "Ohne Betreff";
  return `${timestamp}_${cleanSubject}${suffix}.eml`;
}

function getMailSyncErrorMessage(error) {
  if (error instanceof GraphApiError && error.status === 404) {
    return error.operation === "SharePoint-Upload"
      ? "Der SharePoint-Zielordner /E-Mails fehlt oder ist während des Uploads nicht mehr verfügbar."
      : "Die Nachricht oder SharePoint-Ressource wurde nicht gefunden.";
  }
  return "Microsoft Graph konnte die E-Mail nicht synchronisieren.";
}

app.http("sync-email", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "sync-email",
  handler: async (request, context) => {
    try {
      let body;
      try {
        body = await request.json();
      } catch {
        throw new RequestError(400, "Ungültiger JSON-Request.");
      }

      const payload = parsePayload(body);
      const callerEmail = await validateOfficeToken(request, context);
      if (callerEmail !== payload.userEmail.toLowerCase()) {
        throw new RequestError(403, "Die angegebene Mailbox stimmt nicht mit der Outlook-Anmeldung überein.");
      }

      const accessToken = await getGraphAccessToken();
      const site = await resolveSharePointSite(payload.teamspaceId, accessToken);
      await verifyMailFolder(site.id, accessToken);
      const items = [];
      let singleFailureStatus = 500;

      for (let index = 0; index < payload.itemIds.length; index += 1) {
        const subject = payload.subjects[index] || "";
        const suffix = payload.itemIds.length > 1 ? `_${index + 1}` : "";

        try {
          const messageUrl = `${GRAPH_BASE_URL}/users/${encodeURIComponent(payload.userEmail)}/messages/${encodeURIComponent(payload.itemIds[index])}/$value`;
          const mimeResponse = await graphRequest(messageUrl, accessToken, "E-Mail-Abruf", {
            headers: { Accept: "message/rfc822" },
          });
          if (!mimeResponse.body) throw new Error("Microsoft Graph returned an empty MIME stream.");

          const fileName = createFileName(subject, suffix);
          const uploadUrl = new URL(
            `${GRAPH_BASE_URL}/sites/${encodeURIComponent(site.id)}/drive/root:/E-Mails/${encodeURIComponent(fileName)}:/content`
          );
          await graphRequest(uploadUrl, accessToken, "SharePoint-Upload", {
            method: "PUT",
            headers: { "Content-Type": "message/rfc822" },
            body: mimeResponse.body,
            duplex: "half",
          });
          items.push({ subject, status: "succeeded", fileName });
        } catch (error) {
          if (error instanceof GraphApiError) {
            context.error(`Microsoft Graph ${error.operation} failed with HTTP ${error.status}.`);
            singleFailureStatus = error.status === 404 ? 404 : 502;
          } else {
            context.error("E-Mail-Synchronisierung fehlgeschlagen.", error);
            singleFailureStatus = 500;
          }
          items.push({ subject, status: "failed", error: getMailSyncErrorMessage(error) });
        }
      }

      const failedCount = items.filter((item) => item.status === "failed").length;
      const succeededCount = items.length - failedCount;
      const message = failedCount === 0
        ? succeededCount === 1
          ? "E-Mail erfolgreich nach SharePoint synchronisiert."
          : `${succeededCount} E-Mails erfolgreich nach SharePoint synchronisiert.`
        : `${succeededCount} von ${items.length} E-Mails synchronisiert; ${failedCount} fehlgeschlagen.`;
      const responseStatus = items.length === 1 && failedCount === 1 ? singleFailureStatus : 200;
      const responseBody = { message, items };
      if (items.length === 1 && items[0].status === "succeeded") {
        responseBody.fileName = items[0].fileName;
      } else if (responseStatus !== 200) {
        responseBody.error = items[0].error;
      }
      return jsonResponse(responseStatus, responseBody);
    } catch (error) {
      if (error instanceof RequestError) return jsonResponse(error.status, { error: error.message });

      if (error instanceof GraphApiError) {
        context.error(`Microsoft Graph ${error.operation} failed with HTTP ${error.status}.`);
        if (error.status === 404 && error.operation === "SharePoint-Upload") {
          return jsonResponse(404, {
            error: "Der SharePoint-Zielordner /E-Mails fehlt oder ist während des Uploads nicht mehr verfügbar.",
          });
        }
        return jsonResponse(error.status === 404 ? 404 : 502, {
          error: error.status === 404
            ? "Die Nachricht oder SharePoint-Ressource wurde nicht gefunden."
            : "Microsoft Graph konnte die E-Mail nicht synchronisieren.",
        });
      }

      context.error("E-Mail-Synchronisierung fehlgeschlagen.", error);
      return jsonResponse(500, { error: "E-Mail konnte nicht synchronisiert werden." });
    }
  },
});