# Private GitHub Setup Notes

These notes are for posting CFI DataClean to a private GitHub repository.

## Recommended Safety Check

Before uploading, confirm with DOI/BIA IT or records/security staff that GitHub private repositories are approved for this project.

Do not upload:

- CFI Access databases (`.mdb`, `.accdb`)
- Run exports or run logs
- API keys, passwords, certificates, or `.env` files
- Project-specific manuals or extracted manual text unless approved
- Codex thread backups unless approved

The project includes a `.gitignore` in the main CFI DataClean folder to help keep generated files, Access databases, secrets, manuals, extracted manual text, spreadsheets, reference workbooks, and large source-reference files out of Git.

## Suggested Private Repo Contents

At minimum, the private repo should include:

- `.gitignore`
- `CFIDataClean.cmd`
- `CFIDataClean-How-To-Guide.html`
- `_CFIDataClean_AppFiles\ForestInventoryCleaner.ps1`
- `_CFIDataClean_AppFiles\CFIDataClean.vbs`
- `_CFIDataClean_AppFiles\CFIDataClean-debug.cmd`
- `_CFIDataClean_AppFiles\CFIDataClean-broom.ico`
- `_CFIDataClean_AppFiles\CFIDataClean-tree.ico`
- `_CFIDataClean_AppFiles\README.md`
- `_CFIDataClean_AppFiles\SECURITY.md`

Do not include `_CFIDataClean_AppFiles\field-manual-text.json`, `_CFIDataClean_AppFiles\References`, source manuals, extracted manual text, cleaning spreadsheets, database files, or run exports unless DOI/BIA approves those files for the repository.

## Easy Upload Method

If Git is not installed, use GitHub Desktop:

1. Create a new **private** repository in GitHub.
2. Install GitHub Desktop if approved by IT.
3. In GitHub Desktop, choose **File > Add local repository**.
4. Select the `CFI-DataClean-v1.0.00` folder.
5. Review the changed files before committing.
6. Confirm ignored files are not listed.
7. Commit with a message such as `Initial private CFI DataClean v1.0.00 project`.
8. Publish or push to the private GitHub repository.

## Command Line Method

If Git is installed and available in PowerShell:

```powershell
cd "C:\Users\christopher.lacroix\Documents\Codex\2026-05-28\CFI-DataClean-v1.0.00"
git init
git status
git add .
git status
git commit -m "Initial private CFI DataClean v1.0.00 project"
git branch -M main
git remote add origin https://github.com/YOUR-ACCOUNT/YOUR-PRIVATE-REPO.git
git push -u origin main
```

Use GitHub Desktop or GitHub's web UI to create the remote repository as **Private** before pushing.

## Current Local Note

On this workstation, `git` and `gh` were not found on the command path during setup, so Codex could not push the repository directly.
