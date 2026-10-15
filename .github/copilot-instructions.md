# Repository Guidelines & Architecture: SWE Outlook Project Selector

## Architecture & Deployment
- Runtime: Office.js React Add-in hosted on Azure Static Web Apps.
- CI/CD: Pushes to the `main` branch trigger a GitHub Actions workflow that builds and deploys to Azure.
- Client Testing: Outlook clients (Desktop & Web) load the live application bundle from Azure Static Web Apps, NOT from localhost:3000. Do not start or test against a local dev server unless explicitly requested.

## Platform Constraints (New Outlook & OWA)
- The organization uses New Outlook for Windows (web-engine based).
- Compose Mode Limitation: `item.categories.addAsync` and `mailbox.masterCategories` are strictly blocked by Microsoft during email composition ("Operation not supported"). Do not re-enable category APIs in Compose mode.
- Metadata Storage: Project associations are stored using `Office.context.mailbox.item.loadCustomPropertiesAsync` (`teamspace_project_id`, `teamspace_project_name`, `teamspace_tagged_at`).

## Git & Execution Conventions
- When asked to commit or deploy changes, provide standard git commands (`git add`, `git commit -m "..."`, `git push origin main`).
- Keep `.gitignore` respected: never commit `.zip` files, Azurite databases, or build folders (`dist/`, `package/`).
