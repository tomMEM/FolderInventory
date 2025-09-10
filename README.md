# Inventory Dashboard Application README

## Overview
This guide explains how to package **InventoryDashboard.py** into an executable (`InventoryDashboard.exe`) using **PyInstaller** in a Conda environment defined by `fileinventory.yml`. The app scans folder structures, tracks file changes, and exports results to Excel with filtering options.

---

### 🛠️ Setup Instructions

#### 1. Environment Preparation
- Create a Conda environment using the provided `fileinventory.yml`:
  ```bash
  conda env create -f fileinventory.yml
  conda activate executable_env
  ```
  *Ensure the YAML file is in your project root. It should include dependencies like `pyinstaller`, `pandas`, and `gradio`.*

#### 2. PyInstaller Configuration
- Clean old builds before compiling:
  ```bash
  rm -rf dist build  # Linux/macOS
  rmdir /s /q dist build  # Windows
  ```
- Build the executable:
  ```bash
  pyinstaller InventoryDashboard_cgp.spec
  or
  pyinstaller --clean InventoryDashboard_cgp.spec
  ```

---

### 🚀 Execution

#### Start the App
- Run from the `dist` folder or copy to desktop:
  ```bash
  InventoryDashboard.exe
  ```

#### Key Features
- **Folder Analysis**: Scans directory structures and exports Excel reports
- **Change Tracking**: Identifies:
  - New files added
  - Existing files modified
  - Files removed from previous scans
- **Filter Options**:
  - `Folder:old` - Filter by folder name
  - `IncFolder:did` - Include specific directories
  - `docx` - Filter by file extension

---

### ⚠️ Troubleshooting

#### Gradio Server Conflict
1. Check if port localhost:`Port` is occupied:
   ```bash
   netstat -ano | findstr ":Port"
   ```
2. Terminate conflicting process:
   ```bash
   taskkill /F /PID <PID_NUMBER>
   ```

---

