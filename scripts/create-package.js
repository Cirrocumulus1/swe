const fs = require("node:fs");
const path = require("node:path");
const archiver = require("archiver");
const productionDefaults = require("./production-config");

const rootDirectory = path.resolve(__dirname, "..");
const distDirectory = path.join(rootDirectory, "dist");
const manifestPath = path.join(distDirectory, "manifest.json");
const manifest = JSON.parse(fs.readFileSync(manifestPath, "utf8"));
const appIdUri = process.env.ADDIN_APP_ID_URI?.trim() || productionDefaults.ADDIN_APP_ID_URI;
if (!manifest.webApplicationInfo) {
  throw new Error("Cannot package; webApplicationInfo is missing from the production manifest.");
}
manifest.webApplicationInfo.resource = appIdUri;
fs.writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`);
const packagePath = path.join(
  rootDirectory,
  `swe-outlook-project-selector-v${manifest.version}.zip`
);
const packageFiles = [
  "manifest.json",
  "assets/color.png",
  "assets/outline.png",
  "assets/icon-16.png",
  "assets/icon-32.png",
  "assets/icon-80.png",
];

async function createPackage() {
  for (const relativePath of packageFiles) {
    if (!fs.existsSync(path.join(distDirectory, relativePath))) {
      throw new Error(`Cannot package; required file is missing from dist: ${relativePath}`);
    }
  }

  const output = fs.createWriteStream(packagePath);
  const archive = archiver("zip", { zlib: { level: 9 } });
  const completion = new Promise((resolve, reject) => {
    output.once("close", resolve);
    output.once("error", reject);
    archive.once("error", reject);
  });

  archive.pipe(output);
  for (const relativePath of packageFiles) {
    archive.file(path.join(distDirectory, relativePath), { name: relativePath });
  }

  await archive.finalize();
  await completion;
  console.log(`Created ${path.relative(rootDirectory, packagePath)} (${archive.pointer()} bytes).`);
}

createPackage().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});