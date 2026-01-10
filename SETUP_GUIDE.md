# Distribution Setup Guide

This guide walks you through setting up automated builds for Windows and macOS.

## Step 1: Create Your GitHub Repository

1. Create a new repository on GitHub (e.g., `reference-checker`)
2. Push this project to your repository:

```bash
cd reference-checker
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/reference-checker.git
git push -u origin main
```

## Step 2: Update Configuration

Edit `reference_checker.py` and update the `GITHUB_REPO` constant with your actual GitHub username:

```python
GITHUB_REPO = "YOUR_USERNAME/reference-checker"  # Change this!
```

## Step 3: Configure GitHub Secrets for macOS Signing

Go to your repository on GitHub → Settings → Secrets and variables → Actions → New repository secret

### Required Secrets for macOS

| Secret Name | Description | How to Get It |
|-------------|-------------|---------------|
| `APPLE_CERTIFICATE_BASE64` | Your signing certificate | See below |
| `APPLE_CERTIFICATE_PASSWORD` | Password for the .p12 file | You set this when exporting |
| `APPLE_TEAM_ID` | Your 10-character Team ID | developer.apple.com → Membership |
| `APPLE_ID` | Your Apple ID email | Your Apple Developer account email |
| `APPLE_ID_PASSWORD` | App-specific password | See below |

### Getting Your Signing Certificate

1. **Open Keychain Access** on your Mac
2. Find your certificate: `Developer ID Application: Your Name (XXXXXXXXXX)`
   - If you don't have one, create it at developer.apple.com → Certificates
3. Right-click → **Export** → Save as `Certificates.p12`
4. Set a strong password when prompted (save this!)
5. Convert to base64:
   ```bash
   base64 -i Certificates.p12 | pbcopy
   ```
6. Paste this as `APPLE_CERTIFICATE_BASE64` secret

### Creating App-Specific Password

1. Go to [appleid.apple.com](https://appleid.apple.com)
2. Sign in → Security → App-Specific Passwords
3. Click "Generate an app-specific password"
4. Name it "GitHub Notarization"
5. Copy the generated password
6. Add it as `APPLE_ID_PASSWORD` secret

## Step 4: Windows Build (No Secrets Needed!)

Windows builds work automatically without any secrets. The workflow uses PyInstaller to create a standalone .exe file.

**Note**: Windows users may see a SmartScreen warning the first few times they download. This goes away as more users download the app. Paid code signing ($200-500/year) is optional.

## Step 5: Creating a Release

To trigger automatic builds:

1. Update the `VERSION` in `reference_checker.py`
2. Commit and push your changes
3. Create a tag:
   ```bash
   git tag v1.0.0
   git push origin v1.0.0
   ```
4. GitHub Actions will automatically:
   - Build Windows .exe
   - Build, sign, and notarize macOS .dmg
   - Attach both to a new GitHub Release

## Step 6: Test Your Builds

You can manually trigger builds without creating a release:

1. Go to your repo → Actions
2. Select "Build Windows Executable" or "Build macOS DMG"
3. Click "Run workflow"
4. Download the artifact when complete

## Troubleshooting

### Windows Build Fails
- Check that all dependencies are in `requirements.txt`
- Ensure no syntax errors in the Python code

### macOS Signing Fails
- Verify your certificate is "Developer ID Application" (not "Developer ID Installer")
- Make sure the certificate hasn't expired
- Check that Team ID matches your certificate

### Notarization Fails
- App-specific password might be wrong
- Apple ID might have 2FA issues
- Check Apple's system status: https://developer.apple.com/system-status/

### "Hardened Runtime" Errors
The workflow includes entitlements that allow Python to work with Apple's hardened runtime. If you add native libraries, you may need to adjust the entitlements.

## Files Overview

```
reference-checker/
├── reference_checker.py      # Main application
├── requirements.txt          # Python dependencies
├── README.md                 # User-facing documentation
├── assets/
│   ├── panther_HQ.png       # GUI logo
│   ├── panther_icon.ico     # Windows icon
│   └── panther_icon.icns    # macOS icon
├── docs/
│   └── user_guide.pdf       # User guide (add when ready)
└── .github/
    └── workflows/
        ├── build-windows.yml # Windows build workflow
        └── build-macos.yml   # macOS build workflow
```

## Adding Your User Guide

1. Create your user guide with screenshots
2. Export as PDF named `user_guide.pdf`
3. Place in the `docs/` folder
4. Commit and push
5. It will be bundled with the next release

Users can access it via Help → User Guide in the application.
