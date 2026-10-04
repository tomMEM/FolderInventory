# Inventory Dashboard Application README

## Overview
This guide explains how to package **InventoryDashboard.py** into an executable (`InventoryDashboard.exe`) using **PyInstaller** in a Conda environment defined by `fileinventory.yml`. The app scans[...]

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

## 📄 LastWriteFiles.ipynb - Winword DOCX Assembly Lineage Analyzer

### Overview
**LastWriteFiles.ipynb** is a sophisticated Jupyter notebook that analyzes the creation and modification lineage of Microsoft Word (DOCX) documents. It reconstructs the assembly history of complex documents by detecting content relationships, text reuse patterns, and document evolution over time.

### Key Capabilities

#### 1. **DOCX File Discovery & Metadata Collection**
- Recursively scans a folder and its subfolders for all `.docx` files
- Extracts comprehensive metadata from each document:
  - Creation and modification timestamps
  - Author and last modifier information
  - Revision counts
  - File size and binary fingerprints (SHA256 hashing)
  - Complete document content with paragraph-level granularity

#### 2. **Content Analysis & Text Extraction**
- Extracts current visible text from document bodies, tables, and text boxes
- Normalizes text for lineage comparison (lowercase, punctuation-insensitive)
- Preserves detailed text snapshots for diff analysis
- Generates paragraph hashes and n-gram shingles for similarity detection
- Captures tracked changes, comments, and revision history indicators

#### 3. **Graph Generation: Nodes & Edges**
Produces two CSV files that model document relationships as a directed acyclic graph:

- **`{RUN_NAME}_nodes.csv`**: Lists all discovered DOCX files with extracted metadata
  - File identifiers, paths, creation/modification times
  - Word counts, content hashes, metadata
  - Extraction status and error tracking

- **`{RUN_NAME}_edges.csv`**: Maps parent-child relationships showing document evolution
  - Similarity scores quantifying content reuse (0.0–1.0)
  - Relationship classifications:
    - "Same normalized text / renamed copy" (identical content, different filename)
    - "Revision" (significant overlap in both directions)
    - "Expanded / copied into" (child larger, contains parent)
    - "Excerpt / extracted from" (child smaller, extracted from parent)
    - "Partial reuse" (moderate overlap)
  - Detailed diff metrics when applicable

#### 4. **Advanced Lineage Detection**
- **Content-Based Matching**: Compares normalized paragraph text and word-level n-grams to find genuine predecessors
- **Enhanced Mode**: Explores uncertain candidates when content alone is inconclusive:
  - Filename family matching (e.g., "Report_v1" → "Report_v2")
  - Version progression tracking (v1 → v2 → v3)
  - Chronological analysis (expected modification order)
  - Shingle-based token similarity
  - Metadata priority scoring for plausible candidates
- **Configurable Thresholds**: Adjustable parameters for:
  - Minimum content similarity threshold (default: 0.42)
  - Maximum parent documents per file (default: 3)
  - Time window for parent relationships (default: 90 days)

#### 5. **Text Cache & Future Queries**
- Writes `{RUN_NAME}_text_cache.json.gz` — a compressed snapshot of normalized document content
- Enables fast "target-passage searches" without re-scanning documents
- Stores paragraph locations, normalized words, and fingerprints for efficient lookup

#### 6. **Interactive Dashboard**
- Opens a Gradio-based dashboard interface for visual exploration
- **Visualize Document Assembly**:
  - Browse the document graph in a tree or network view
  - Click on documents to see metadata, content summaries, and relationships
  - Track how documents evolved and where content was reused or extracted
- **Detailed Diff Viewer**: Compare pairs of documents with token-level and paragraph-level differences
- **Timeline View**: See documents ordered by modification time with lineage connections

### How It Works

1. **Discovery Phase**: Scans the specified folder recursively for all `.docx` files
2. **Extraction Phase**: Reads each DOCX (ZIP archive) and extracts text, metadata, and structural features
3. **Comparison Phase**: Computes similarity between all pairs of documents within the time window
4. **Lineage Resolution Phase**:
   - Selects the most likely parent(s) for each document using content and metadata scoring
   - In enhanced mode, flags uncertain candidates for manual review
5. **Output Generation**:
   - Writes CSV files (_nodes.csv, _edges.csv) for import into graph databases or visualization tools
   - Optionally generates a DOT file for Graphviz visualization
   - Creates a text cache for fast passage searches
6. **Dashboard Launch**: Opens an interactive interface to explore the discovered lineage

### Configuration Example

Key parameters in the notebook:

```python
ROOT_DIR = r"Z:\TB"               # Folder to scan
RUN_NAME = "Working_File2nd"      # Output prefix
DAYS_AGO = 120                    # Scan recent files from past N days
MIN_EDGE_SCORE = 0.42             # Minimum similarity to link documents
MAX_PARENTS = 3                   # Max parent documents per file
ENHANCED_LINEAGE = True           # Enable uncertain candidate detection
DETAIL_DIFF_THRESHOLD = 0.95      # Score threshold for detailed diff
```

### Use Cases

- **Complex Document Audits**: Track how a multi-version DOCX evolved across revisions
- **Content Reuse Analysis**: Discover which documents were copied, excerpted, or merged
- **Collaboration Tracking**: Identify authorship and modification patterns
- **Compliance & Record Keeping**: Maintain a complete genealogy of document versions
- **Data Quality**: Detect duplicate or near-duplicate documents that may be sources of confusion

### Output Files

After running, you'll find:

- **Working_File2nd_nodes.csv** — Document catalog
- **Working_File2nd_edges.csv** — Lineage relationships with similarity scores
- **Working_File2nd_text_cache.json.gz** — Compressed text snapshot for fast searches
- **Working_File2nd_history.dot** — Graphviz format for graph visualization
- **Dashboard (interactive)** — Gradio web interface for exploration
