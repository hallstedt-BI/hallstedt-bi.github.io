# Excel to GitHub Pages Automation

This setup allows you to export data from a SharePoint Excel file directly to your GitHub Pages site using Office Scripts.

## How It Works

1. **Office Script** in your SharePoint Excel file extracts data from the "Tabell HT25" sheet
2. **GitHub API** receives the data and creates/updates `data.json` in your repository
3. **GitHub Pages** automatically rebuilds and displays the updated data as a table

## Setup Instructions

### Step 1: Create GitHub Personal Access Token

1. Go to **GitHub.com** → **Settings** → **Developer settings** → **Personal access tokens** → **Tokens (classic)**
2. Click **"Generate new token (classic)"**
3. Give it a name like "Excel Office Script"
4. Set expiration (recommend 1 year)
5. Check: ✅ **repo** (Full control of private repositories)
6. Click **"Generate token"** and **copy the token** (you'll only see it once!)

### Step 2: Add Office Script to Your Excel File

1. Open your SharePoint Excel file in Excel Online
2. Go to **Automate** tab → **New Script**
3. Replace the default code with the contents of `office-script.ts`
4. **Update the configuration** at the top:
   ```typescript
   const GITHUB_TOKEN = 'your_token_here';  // Paste your GitHub token
   const GITHUB_REPO = 'your-username/your-repo-name';  // Update with your repo
   ```
5. Save the script with a name like "Export to GitHub"

### Step 3: Run the Script

1. In Excel, go to **Automate** tab
2. Click your "Export to GitHub" script
3. Click **Run**
4. Check the output panel for success/error messages
5. Your GitHub Pages site should update within a few minutes!

## Files

- `index.html` - The GitHub Pages site that displays your data
- `office-script.ts` - The Office Script code to copy into Excel
- `data.json` - Generated automatically by the Office Script

## Features

- ✅ No authentication required for website visitors
- ✅ Automatic table generation from Excel data  
- ✅ Shows last updated timestamp
- ✅ Responsive design for mobile/desktop
- ✅ Auto-refreshes every 5 minutes
- ✅ One-click manual trigger from Excel

## Troubleshooting

**Script fails with 401 error**: Check your GitHub token permissions
**Script fails with 404 error**: Verify your repository name is correct
**No data appears**: Ensure the "Tabell HT25" sheet exists in your Excel file
**Website shows loading**: Check if `data.json` was created in your repository
