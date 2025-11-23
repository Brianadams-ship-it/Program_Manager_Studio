# Optimum Upgrade – Digital PM, Verification & Defense Toolkits

This repository contains the full Optimum Upgrade product suite, including:

- **Digital PM Toolkits** (WBS, Compliance Matrix, Risk Register, Cost Dashboard)
- **Verification & Test Toolkits** (Test Case Matrix, Verification Report Generator, Integration Matrix)
- **MIL-SPEC & DID Knowledge Base**
- **MIL-SPEC & DID Explorer (HTML App)**
- **DID Template Library (DOCX)**
- **SBIR Proposal Prompt Toolkit**
- **Product pages, bundle pages, Gumroad copy**
- **Branding graphics, cover images, hero banners**
- **Excel, DOCX, PDF, and HTML assets**
- **Site styling (CSS)**

All tools are fully offline and can be used on:
- Windows  
- macOS  
- Linux  
- SharePoint / OneDrive  
- GitHub Pages  
- Local browsers  
- Secure government systems (no external calls)

---

## 📁 Folder Structure

```
/pages/         → All product pages, apps, bundles, Gumroad copy, index
/images/        → Covers, branding, promo graphics, hero images
/docs/          → DID templates, PDFs, Word docs
/toolkits/      → Excel-based PM & verification tools
/style/         → Global CSS
```

---

## 🌐 GitHub Pages Deployment

1. Upload all files/folders in this repo to your GitHub repository root.
2. Go to **Settings → Pages**
3. Set:
   - Source: `main`
   - Folder: `/ (root)`
4. Your live site will appear at:

```
https://<username>.github.io/<repo-name>/pages/index.html
```

---

## 📦 ZIP Packaging Script (Optional)

### Windows PowerShell

```powershell
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$zipPath = Join-Path $root "OptimumUpgrade_ALL_latest.zip"
if (Test-Path $zipPath) { Remove-Item $zipPath }
Add-Type -AssemblyName 'System.IO.Compression.FileSystem'
[System.IO.Compression.ZipFile]::CreateFromDirectory($root, $zipPath)
Write-Host "Created ZIP: $zipPath"
```

### macOS/Linux

```bash
zip -r OptimumUpgrade_ALL_latest.zip . \
    -x "*.git*" \
    -x "*.DS_Store" \
    -x "OptimumUpgrade_ALL_latest.zip"
```

---

## 📧 Contact / Support

For customization, consulting, or building advanced PM/verification systems:  
**Optimum Upgrade – AI-Powered Engineering & Program Tools**
