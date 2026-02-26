# GitHub Free Hosting (GitHub Pages)

This folder is prepared for free hosting via GitHub Pages.

## 1) Create a new GitHub repo
- Name suggestion: `it-asset-tracker`
- Keep it Public (free Pages is simplest with Public repo)

## 2) Push this copied folder
Run these commands inside this folder:

```bash
git init
git add .
git commit -m "Initial IT asset tracker"
git branch -M main
git remote add origin https://github.com/<your-username>/<your-repo>.git
git push -u origin main
```

## 3) Enable Pages in GitHub
- Open repo in GitHub
- Go to `Settings` -> `Pages`
- Under Source, choose `GitHub Actions`

## 4) Wait for deployment
- Go to `Actions` tab
- Wait for workflow `Deploy static site to GitHub Pages` to finish
- Your app URL will be:
  `https://<your-username>.github.io/<your-repo>/`

## Notes
- Your data is stored in browser localStorage, so data is browser/device specific.
- Use your backup/export buttons to move data between devices.




