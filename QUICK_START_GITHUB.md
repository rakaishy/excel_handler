# Quick Start - GitHub Actions Build

## TL;DR - Get Windows .exe in 3 Steps

```bash
# 1. Push to GitHub
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/YOUR_USERNAME/excel_handler.git
git push -u origin main

# 2. Create a release tag
git tag v1.0.0
git push origin v1.0.0

# 3. Wait 5 minutes, then download from:
# https://github.com/YOUR_USERNAME/excel_handler/releases
```

## What Happens

1. **You push a tag** (`v1.0.0`) from your Linux machine
2. **GitHub Actions automatically**:
   - Spins up a Windows VM
   - Installs Python
   - Installs dependencies
   - Builds `Excel_Handler.exe`
   - Creates a GitHub Release
3. **Users download** the `.exe` from GitHub Releases
4. **Users run** by double-clicking - no installation needed!

## First Time Setup

### 1. Create GitHub Repository

1. Go to https://github.com/new
2. Name it: `excel_handler`
3. Make it public (for free GitHub Actions)
4. Don't initialize with README (you already have one)
5. Click "Create repository"

### 2. Push Your Code

```bash
cd /home/rakaishy/Documents/repos/CascadeProjects/windsurf-project/excel_handler

# Initialize git (if not already)
git init

# Add all files
git add .

# Commit
git commit -m "Initial commit - Excel Handler with GitHub Actions"

# Add remote (replace YOUR_USERNAME)
git remote add origin https://github.com/YOUR_USERNAME/excel_handler.git

# Push
git push -u origin main
```

### 3. Create Your First Release

```bash
# Create version tag
git tag v1.0.0

# Push the tag
git push origin v1.0.0
```

### 4. Wait for Build

1. Go to your repository on GitHub
2. Click "Actions" tab
3. Watch the build progress (~5 minutes)
4. When complete, go to "Releases"
5. Download `Excel_Handler.exe`

## Sharing with Users

### Simple Instructions for Users

Send them this:

---

**Download Excel Handler:**

1. Click here: `https://github.com/YOUR_USERNAME/excel_handler/releases/latest`
2. Download `Excel_Handler.exe`
3. Double-click to run

---

### Direct Download Link

```
https://github.com/YOUR_USERNAME/excel_handler/releases/latest/download/Excel_Handler.exe
```

Users can click this link to download directly!

## Making Updates

When you fix bugs or add features:

```bash
# Make changes
nano src/main.py

# Commit
git add .
git commit -m "Fixed validation bug"
git push origin main

# Create new version
git tag v1.0.1
git push origin v1.0.1
```

GitHub Actions will automatically build the new version!

## Manual Trigger (Without Tags)

If you just want to test the build:

1. Go to your repository on GitHub
2. Click "Actions" tab
3. Click "Build Windows Executable"
4. Click "Run workflow" dropdown
5. Click green "Run workflow" button
6. Download from "Artifacts" section (not a release)

## Cost

- **Public repository**: FREE (unlimited builds)
- **Private repository**: 2,000 free minutes/month

Each build takes ~5 minutes, so you get ~400 free builds/month on private repos.

## Troubleshooting

### "remote: Repository not found"

Your repository URL is wrong. Check:
```bash
git remote -v
```

Update if needed:
```bash
git remote set-url origin https://github.com/YOUR_USERNAME/excel_handler.git
```

### Build Failed

1. Go to "Actions" tab
2. Click the failed workflow
3. Check error logs
4. Fix the issue
5. Push again: `git push origin main`

### No Release Created

Make sure:
- Tag starts with `v` (e.g., `v1.0.0`)
- Repository is public or you have Actions enabled
- Build completed successfully

## Benefits

✅ Build Windows .exe from Linux
✅ No Windows machine needed
✅ Automated and free
✅ Professional release management
✅ Easy for users to download
✅ Version control built-in

## See Also

- `GITHUB_ACTIONS_GUIDE.md` - Detailed guide
- `.github/workflows/build-windows.yml` - The workflow file
- `README.md` - User documentation
