import jwt from "jsonwebtoken";
import jwksClient from "jwks-rsa";

const jwksClientsByTenant = new Map();

function getJwksClient(tenantId) {
  let client = jwksClientsByTenant.get(tenantId);
  if (!client) {
    const jwksUri = `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/discovery/v2.0/keys`;
    client = jwksClient({
      jwksUri,
      cache: true,
      cacheMaxEntries: 5,
      cacheMaxAge: 10 * 60 * 1000,
      rateLimit: true,
      jwksRequestsPerMinute: 10,
      timeout: 5000,
    });
    jwksClientsByTenant.set(tenantId, client);
  }
  return client;
}

function parseTokenHeader(token) {
  const segments = token.split(".");
  if (segments.length !== 3 || !/^[A-Za-z0-9_-]+$/.test(segments[0])) {
    const error = new Error("The Outlook token is not a three-part JWT.");
    error.code = "ERR_JWT_HEADER_INVALID";
    throw error;
  }

  let header;
  try {
    header = JSON.parse(Buffer.from(segments[0], "base64url").toString("utf8"));
  } catch {
    const error = new Error("The Outlook token header is not valid base64url JSON.");
    error.code = "ERR_JWT_HEADER_INVALID";
    throw error;
  }
  if (!header || typeof header !== "object" || Array.isArray(header)) {
    const error = new Error("The Outlook token header must be a JSON object.");
    error.code = "ERR_JWT_HEADER_INVALID";
    throw error;
  }
  return header;
}

export async function verifyOfficeToken(token, { tenantId, clientId, appIdUri }, createClient = getJwksClient) {
  const header = parseTokenHeader(token);
  const hasKeyId = typeof header.kid === "string" && header.kid.length > 0;
  if (header.alg !== "RS256" || !hasKeyId) {
    const algorithm = typeof header.alg === "string" ? header.alg : "missing";
    const error = new Error(`Expected RS256 with a key ID; received alg=${algorithm}, kidPresent=${hasKeyId}.`);
    error.code = "ERR_JWT_HEADER_INVALID";
    throw error;
  }

  const signingKey = await createClient(tenantId).getSigningKey(header.kid);
  if (signingKey.alg && signingKey.alg !== "RS256") {
    const error = new Error("The signing key algorithm does not match RS256.");
    error.code = "ERR_JWT_ALG_MISMATCH";
    throw error;
  }

  const payload = jwt.verify(token, signingKey.getPublicKey(), {
    algorithms: ["RS256"],
    issuer: [
      `https://login.microsoftonline.com/${tenantId}/v2.0`,
      `https://sts.windows.net/${tenantId}/`,
    ],
    audience: [clientId, appIdUri, `api://${clientId}`],
  });

  if (!payload || typeof payload !== "object") {
    const error = new Error("The Outlook token payload must be a JSON object.");
    error.code = "ERR_JWT_PAYLOAD_INVALID";
    throw error;
  }

  return payload;
}