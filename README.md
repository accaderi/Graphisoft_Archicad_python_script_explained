# Archicad Automation Scripts

This repository contains Python scripts designed to automate various tasks in Archicad using its API. Each script addresses a specific functionality, ranging from managing project elements to generating reports and handling conflicts.

## Scripts Overview

### 1. **Unused Items in View Map**
- **File:** `unused_items_in_view_map_explained.py`
- **Purpose:** Identifies unused items in the View Map and organizes them into folders.
- **Features:**
  - Moves unused navigator items to a designated folder.
  - Renames folders from previous runs for better organization.
  - Ensures only unused "parent" items are included in the list.

---

### 2. **Zone Numbering**
- **File:** `zone_numbering_explained.py`
- **Purpose:** Automates numbering of zones based on their position and level.
- **Features:**
  - Groups zones by levels and sides of the building.
  - Assigns unique numbers to zones using a predefined format.
  - Handles tolerance limits for grouping zones.

---

### 3. **Element ID Conflict Detection**
- **File:** `elementID_conflict_explained.py`
- **Purpose:** Detects conflicts in Element IDs within the project.
- **Features:**
  - Identifies duplicate Element IDs across all elements.
  - Outputs detailed conflict messages for resolution.
  - Confirms when no conflicts are found.

---

### 4. **Zone Overall Dimensions**
- **File:** `zone_overall_dimensions_explained.py`
- **Purpose:** Calculates and assigns overall dimensions (width x height) for zones.
- **Features:**
  - Determines bounding box dimensions for each zone.
  - Formats dimensions with the larger value first (office preference).
  - Updates zone properties with calculated values.

---

### 5. **Room Report Generator**
- **File:** `room_report_explained.py`
- **Purpose:** Generates detailed Excel reports for rooms in the project.
- **Features:**
  - Extracts room properties like name, number, category, area, volume, etc.
  - Includes adjacent zones, equipment details, and openings in the report.
  - Uses a predefined template for structured output.

---

### 6. **Chair Numbering**
- **File:** `chair_numbering__explained.py`
- **Purpose:** Automates numbering of chairs in an auditorium based on layout.
- **Features:**
  - Groups chairs by rows and sides (left/right).
  - Assigns unique IDs using a row-index format (e.g., `A.1/Right`).
  - Handles tolerance limits for grouping chairs.

---

### 7. **Parking Space Numbering**
- **File:** `parking_spaces_explained.py`
- **Purpose:** Automates numbering of parking spaces based on their layout.
- **Features:**
  - Groups parking spaces by levels and rows.
  - Assigns unique IDs using a predefined format (e.g., `P112`).
  - Handles tolerance limits for grouping spaces.

---

### 8. **Excel Export Utility**
- **File:** `excel_export_explained.py`
- **Purpose:** Exports element properties to an Excel file for beams and walls.
- **Features:**
  - Extracts properties like Element ID, Height, Width, Thickness, etc.
  - Creates separate worksheets for beams and walls.
  - Auto-adjusts column widths for readability.

---

### 9. **Excel Import Utility**
- **File:** `excel_import_explained.py`
- **Purpose:** Imports property values from an Excel file into Archicad elements.
- **Features:**
  - Reads element IDs and property values from Excel sheets.
  - Updates corresponding element properties in Archicad.
  - Verifies changes by printing updated values to the console.


## Requirements
1. Archicad software must be open with an active project file (`.pln`).
2. Python environment with necessary dependencies installed:
   - `archicad` API module
   - `openpyxl` (for Excel operations)


## How to Use
1. Clone this repository to your local machine:
    ```
    git clone <repository-url>
    ```
2. Open your Archicad project file (`.pln`).
3. Run the desired script using Python:
    ```
    python <script_name>.py
    ```
4. Follow any prompts or outputs displayed in the console.


## Notes
- Ensure proper configuration of variables within each script before execution (e.g., folder names, output paths).
- Some scripts rely on predefined templates or classification systems; verify their availability before running.
