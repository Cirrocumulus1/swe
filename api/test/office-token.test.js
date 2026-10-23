import assert from "node:assert/strict";
import { generateKeyPairSync } from "node:crypto";
import { once } from "node:events";
import { createServer } from "node:http";
import test from "node:test";
import jwt from "jsonwebtoken";
import jwksClient from "jwks-rsa";
import { verifyOfficeToken } from "../src/office-token.js";

const tenantId = "00000000-0000-4000-8000-000000000001";
const clientId = "00000000-0000-4000-8000-000000000002";
const appIdUri = `api://example.test/${clientId}`;
const issuer = `https://login.microsoftonline.com/${tenantId}/v2.0`;
const keyId = "local-rs256-test-key";

async function createTestJwks() {
  const { privateKey, publicKey } = generateKeyPairSync("rsa", { modulusLength: 2048 });
  const jwk = {
    ...publicKey.export({ format: "jwk" }),
    kid: keyId,
    use: "sig",
    alg: "RS256",
    key_ops: ["verify"],
  };
  const server = createServer((request, response) => {
    if (request.url !== "/keys") {
      response.writeHead(404).end();
      return;
    }
    response.writeHead(200, { "Content-Type": "application/json" });
    response.end(JSON.stringify({ keys: [jwk] }));
  });
  server.listen(0, "127.0.0.1");
  await once(server, "listening");
  const address = server.address();
  const client = jwksClient({
    jwksUri: `http://127.0.0.1:${address.port}/keys`,
    cache: true,
    rateLimit: true,
    timeout: 2000,
  });

  return { client, privateKey, server };
}

function signToken(privateKey, claims, algorithm = "RS256") {
  return jwt.sign(claims, privateKey, { algorithm, keyid: keyId });
}

test("verifies an Entra-style RS256 token by kid against a JWKS endpoint", async (t) => {
  const { client, privateKey, server } = await createTestJwks();
  t.after(() => new Promise((resolve, reject) => server.close((error) => error ? reject(error) : resolve())));

  const token = signToken(privateKey, {
    aud: clientId,
    exp: Math.floor(Date.now() / 1000) + 60,
    iss: issuer,
    scp: "access_as_user",
    preferred_username: "test@example.com",
  });
  const payload = await verifyOfficeToken(token, { tenantId, clientId, appIdUri }, () => client);

  assert.equal(payload.preferred_username, "test@example.com");
  assert.equal(payload.scp, "access_as_user");
});

test("rejects a token with an unexpected audience", async (t) => {
  const { client, privateKey, server } = await createTestJwks();
  t.after(() => new Promise((resolve, reject) => server.close((error) => error ? reject(error) : resolve())));

  const token = signToken(privateKey, {
    aud: "another-client",
    exp: Math.floor(Date.now() / 1000) + 60,
    iss: issuer,
    scp: "access_as_user",
  });

  await assert.rejects(
    verifyOfficeToken(token, { tenantId, clientId, appIdUri }, () => client),
    { name: "JsonWebTokenError", message: /^jwt audience invalid/ }
  );
});

test("rejects a token with an unexpected issuer", async (t) => {
  const { client, privateKey, server } = await createTestJwks();
  t.after(() => new Promise((resolve, reject) => server.close((error) => error ? reject(error) : resolve())));

  const token = signToken(privateKey, {
    aud: clientId,
    exp: Math.floor(Date.now() / 1000) + 60,
    iss: "https://attacker.example/tenant/v2.0",
  });

  await assert.rejects(
    verifyOfficeToken(token, { tenantId, clientId, appIdUri }, () => client),
    { name: "JsonWebTokenError", message: /^jwt issuer invalid/ }
  );
});

test("rejects an expired token", async (t) => {
  const { client, privateKey, server } = await createTestJwks();
  t.after(() => new Promise((resolve, reject) => server.close((error) => error ? reject(error) : resolve())));

  const token = signToken(privateKey, {
    aud: clientId,
    exp: Math.floor(Date.now() / 1000) - 1,
    iss: issuer,
  });

  await assert.rejects(
    verifyOfficeToken(token, { tenantId, clientId, appIdUri }, () => client),
    { name: "TokenExpiredError", message: "jwt expired" }
  );
});

test("rejects non-RS256 algorithms before resolving a key", async () => {
  const token = signToken("local-test-secret", {
    aud: clientId,
    iss: issuer,
  }, "HS256");

  await assert.rejects(
    verifyOfficeToken(token, { tenantId, clientId, appIdUri }, () => {
      throw new Error("JWKS lookup should not run for HS256");
    }),
    { code: "ERR_JWT_HEADER_INVALID" }
  );
});