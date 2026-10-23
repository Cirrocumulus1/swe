const fs = require("node:fs");
const path = require("node:path");
const productionDefaults = require("./production-config");

const distDirectory = path.resolve(__dirname, "..", "dist");
const manifestPath = path.join(distDirectory, "manifest.json");
const manifest = JSON.parse(fs.readFileSync(manifestPath, "utf8"));
const productionBaseUrl = new URL(
  process.env.ADDIN_BASE_URL?.trim() || productionDefaults.ADDIN_BASE_URL
);
const manifestId = process.env.ADDIN_MANIFEST_ID?.trim() || productionDefaults.ADDIN_MANIFEST_ID;
const entraClientId = process.env.ENTRA_CLIENT_ID?.trim() || productionDefaults.ENTRA_CLIENT_ID;
const appIdUri = process.env.ADDIN_APP_ID_URI?.trim() || productionDefaults.ADDIN_APP_ID_URI;

function fail(message) {
  console.error(`Production manifest validation failed: ${message}`);
  process.exit(1);
}

if (manifest.version !== productionDefaults.PRODUCTION_VERSION) {
  fail(`expected version ${productionDefaults.PRODUCTION_VERSION}, got ${manifest.version}.`);
}
if (manifest.id !== manifestId) fail("the package ID doesn't match the production app ID.");
if (manifest.webApplicationInfo?.id !== entraClientId) {
  fail("webApplicationInfo.id doesn't match the production Entra app ID.");
}
if (manifest.webApplicationInfo?.resource !== appIdUri) {
  fail("webApplicationInfo.resource doesn't match the configured Entra app ID URI.");
}
if (
  manifest.name?.short !== productionDefaults.PRODUCTION_NAME_SHORT ||
  manifest.name?.full !== productionDefaults.PRODUCTION_NAME_FULL
) {
  fail("the manifest still contains a non-production app name.");
}
if (JSON.stringify(manifest).toLowerCase().includes("localhost")) {
  fail("a localhost URL remains in the production manifest.");
}

const expectedHost = productionBaseUrl.host;
if (manifest.validDomains?.length !== 1 || manifest.validDomains[0] !== expectedHost) {
  fail(`validDomains must contain only ${expectedHost}.`);
}

const hostedUrls = [];
function collectHostedUrls(value) {
  if (Array.isArray(value)) {
    value.forEach(collectHostedUrls);
    return;
  }
  if (!value || typeof value !== "object") return;

  for (const [key, nestedValue] of Object.entries(value)) {
    if ((key === "page" || key === "url") && typeof nestedValue === "string") {
      hostedUrls.push(nestedValue);
    } else {
      collectHostedUrls(nestedValue);
    }
  }
}

collectHostedUrls(manifest);
for (const value of hostedUrls) {
  let url;
  try {
    url = new URL(value);
  } catch {
    fail(`invalid hosted URL: ${value}.`);
  }
  if (url.host !== expectedHost || url.protocol !== "https:") {
    fail(`hosted URL doesn't use the production HTTPS host: ${value}.`);
  }
}

const requiredAssets = [
  manifest.icons?.color,
  manifest.icons?.outline,
  "assets/icon-16.png",
  "assets/icon-32.png",
  "assets/icon-80.png",
];
for (const relativeAssetPath of requiredAssets) {
  if (typeof relativeAssetPath !== "string" || path.isAbsolute(relativeAssetPath)) {
    fail("a required icon path is missing or invalid.");
  }
  const assetPath = path.resolve(distDirectory, relativeAssetPath);
  if (!assetPath.startsWith(`${distDirectory}${path.sep}`) || !fs.existsSync(assetPath)) {
    fail(`required icon is missing from dist: ${relativeAssetPath}.`);
  }
}

console.log("Production manifest values and required icon assets are valid.");