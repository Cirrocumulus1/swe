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

export async function verifyOfficeToken(token, { tenantId, clientId, appIdUri }, createClient = getJwksClient) {
  const decoded = jwt.decode(token, { complete: true });
  const header = decoded && typeof decoded === "object" ? decoded.header : undefined;
  if (header?.alg !== "RS256" || typeof header.kid !== "string" || !header.kid) {
    const error = new Error("The Outlook token must use RS256 and include a key ID.");
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