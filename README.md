# Bearing Graph - Interactive Pile Driving Analysis Tool

An interactive web-based tool for visualizing and analyzing pile driving analysis results using Dash and Plotly.

## Features

- **Interactive filtering** by Position (MP), Soil Bounds, and Method
- **Real-time graph updates** as you change filters
- **Multiple data combinations**:
  - Soil bounds: lb (lower bound), be (best estimate), ub (upper bound)
  - Methods: MD, AH, MY
  - Soil combinations: lb+be, lb+ub, be+ub
  - Method combinations: MD+AH, MD+MY, AH+MY
- **Clean visualization** with customized line styles and colors

## Installation

### Requirements
- Python >= 3.14
- Virtual environment (recommended)

### Setup

1. **Clone or navigate to the project directory**:
```bash
cd C:\Users\vasav\PyCharmProjects\VASAV-at-work
```

2. **Activate the virtual environment**:
```bash
.\.venv\Scripts\Activate.ps1
```

3. **Install dependencies**:
```bash
pip install -r requirements.txt
```

## Usage

### Run the Dash App

```bash
python HOW03/bearing_graph_dash.py --root "K:\dozr\HOW03\GEO\05_Driveability\20240909_Final Design for Certification\variations\const_en_MENCK_original_cans\monopiles"
```

The app will start on `http://127.0.0.1:8050`

Open this URL in your web browser to access the interactive tool.

### Using the Interface

1. **Position (MP) Dropdown**: Select which wind farm location(s) to view
2. **Soil Bounds Dropdown**: Choose soil bound type(s) (lb, be, ub)
3. **Method Dropdown**: Select analysis method(s) (MD, AH, MY)

All selections are **independent and stateful** - you can change any dropdown at any time and the graph will update correctly.

### Stopping the Server

Press `Ctrl+C` in the terminal to stop the Dash server.

## Plot Details

- **X-axis**: blowcount / 4 (units: bl/25cm)
- **Y-axis**: SRD (units: MN)
- **Line styles**: 
  - Solid line = lb (lower bound)
  - Dotted line = be (best estimate)
  - Dashed line = ub (upper bound)
- **Line colors**:
  - Blue = MD (Monodrive method)
  - Red = AH (Auxiliary hydraulic method)
  - Green = MY (Maynard method)

## Data Format

The tool expects CSV files in the following structure:
- Root folder contains subfolders named after positions (A01, A02, etc.)
- Each subfolder contains a file: `results_PileDrivingAnalysis-<POSITION>.csv`
- CSV files must use semicolon (`;`) as delimiter
- Tables start with `**` markers followed by headers and data rows

## GitHub Workflow

### Initial Setup (First Time Only)

1. **Initialize git in your project** (if not already done):
```bash
cd C:\Users\vasav\PyCharmProjects\VASAV-at-work
git init
```

2. **Add remote repository**:
```bash
git remote add origin https://github.com/<your-username>/<repo-name>.git
```

### Standard Workflow

1. **Check status**:
```bash
git status
```

2. **Stage changes**:
```bash
# Stage specific files
git add HOW03/bearing_graph_dash.py requirements.txt pyproject.toml README.md

# Or stage all changes
git add .
```

3. **Commit changes**:
```bash
git commit -m "Add bearing graph interactive tool with stateful dropdown filtering"
```

4. **Push to GitHub**:
```bash
git push -u origin main
```
(Use `main` or `master` depending on your default branch)

### Example Commit Messages

Good practice:
```bash
# Feature addition
git commit -m "Add bearing_graph_dash.py with stateful filtering"

# Updates
git commit -m "Update dependencies: Dash 3.3.0, Flask 3.1.2 for Python 3.14"

# Cleanup
git commit -m "Remove unused plotting scripts, clean up project"
```

### Viewing History

```bash
# View commit history
git log --oneline

# View detailed changes in last commit
git show

# View differences before committing
git diff
```

## Troubleshooting

### Port 8050 already in use
Change the port in the script or use:
```bash
python HOW03/bearing_graph_dash.py --root "<path>" --port 8051
```

### Data not loading
Ensure the root path exists and contains subfolders with CSV files in the correct format.

### Graph not updating
Refresh the browser or restart the server.

## Project Structure

```
VASAV-at-work/
├── HOW03/
│   ├── bearing_graph_dash.py    # Main interactive Dash app
│   └── extract_results.py        # Data extraction utility
├── HOW04/
├── pyproject.toml               # Project configuration
├── requirements.txt             # Python dependencies
└── README.md                    # This file
```

## License

Internal project - VASAV at work

---

**Last Updated**: November 21, 2025

