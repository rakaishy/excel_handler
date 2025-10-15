# GitHub Actions - Automated Windows Build

This project uses GitHub Actions to automatically build Windows executables, so you don't need a Windows machine!

## How It Works

When you push a version tag (like `v1.0.0`), GitHub Actions will:
1. ✅ Spin up a Windows virtual machine
2. ✅ Install Python and dependencies
3. ✅ Build the Windows executable
4. ✅ Create a GitHub Release with the `.exe` file
5. ✅ Users can download the `.exe` directly from GitHub

## Setup Steps

### 1. Push Your Code to GitHub

```bash
# Initialize git (if not already done)
git init

# Add all files
git add .

# Commit
git commit -m "Initial commit - Excel Handler"

# Add your GitHub repository as remote
git remote add origin https://github.com/YOUR_USERNAME/excel_handler.git

# Push to GitHub
git push -u origin main
```

### 2. Create a Release

**Option A: Using Git Tags (Recommended)**

```bash
# Create a version tag
git tag v1.0.0

# Push the tag to GitHub
git push origin v1.0.0
```

This will automatically trigger the build and create a release!

**Option B: Manual Trigger**

1. Go to your GitHub repository
2. Click on "Actions" tab
3. Click on "Build Windows Executable" workflow
4. Click "Run workflow" button
5. The executable will be available as an artifact

### 3. Download the Executable

After the build completes (takes ~5 minutes):

**If you used a tag:**
1. Go to your repository on GitHub
2. Click "Releases" on the right side
3. Download `Excel_Handler.exe`

**If you manually triggered:**
1. Go to "Actions" tab
2. Click on the completed workflow run
3. Scroll down to "Artifacts"
4. Download `Excel_Handler-Windows`

## Distributing to Users

### Method 1: GitHub Releases (Recommended)

1. Users go to: `https://github.com/YOUR_USERNAME/excel_handler/releases`
2. Click on the latest release
3. Download `Excel_Handler.exe`
4. Double-click to run

### Method 2: Direct Download Link

Share this link with users:
```
https://github.com/YOUR_USERNAME/excel_handler/releases/latest/download/Excel_Handler.exe
```

They can click it to download directly!

## Version Management

### Creating New Versions

When you make changes and want to release a new version:

```bash
# Make your changes
git add .
git commit -m "Fixed validation bug"

# Create new version tag
git tag v1.0.1

# Push code and tag
git push origin main
git push origin v1.0.1
```

GitHub Actions will automatically build and create a new release!

### Version Numbering

Use semantic versioning:
- `v1.0.0` - Major release
- `v1.0.1` - Bug fix
- `v1.1.0` - New feature
- `v2.0.0` - Breaking changes

## Troubleshooting

### Build Failed

1. Go to "Actions" tab on GitHub
2. Click on the failed workflow
3. Check the error logs
4. Common issues:
   - Missing dependencies in `requirements.txt`
   - Syntax errors in `build.py`
   - Import errors in `src/main.py`

### No Release Created

Make sure:
- You pushed a tag starting with `v` (e.g., `v1.0.0`)
- The build completed successfully
- You have releases enabled in repository settings

### Can't Download Executable

Check:
- Repository is public (or user has access if private)
- Release was created successfully
- Executable was uploaded (check release assets)

## Benefits

✅ **No Windows machine needed** - Build on Linux, deploy to Windows
✅ **Automated** - Just push a tag, get an executable
✅ **Free** - GitHub Actions is free for public repositories
✅ **Version control** - All releases are tracked
✅ **Easy distribution** - Users download from GitHub
✅ **Professional** - Proper release management

## Alternative: Manual Trigger

If you don't want to use tags, you can manually trigger builds:

1. Go to GitHub repository
2. Click "Actions" tab
3. Select "Build Windows Executable"
4. Click "Run workflow"
5. Select branch (usually `main`)
6. Click "Run workflow" button

The executable will be available as an artifact (not a release).

## For Non-Technical Users

Share this simple instruction with your users:

---

**How to Download Excel Handler:**

1. Go to: [Your GitHub Releases Page]
2. Click the green "Excel_Handler.exe" link
3. Save the file to your computer
4. Double-click the file to run

That's it! No installation needed.

---

## Cost

- **Public repositories**: 100% FREE
- **Private repositories**: 2,000 free minutes/month (each build takes ~5 minutes)

## Next Steps

1. Push your code to GitHub
2. Create your first release: `git tag v1.0.0 && git push origin v1.0.0`
3. Wait 5 minutes for the build
4. Share the download link with users!
