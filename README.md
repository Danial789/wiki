# Lean Digital Twin Application

A comprehensive application for interacting with GraphDB repositories, visualizing data models, and linking data to Excel datasheets to create a "Lean Digital Twin."

## Features

- **Graphical Model Visualization**: Interactive network visualization of GraphDB relationships
- **GraphDB Integration**: Connect to SPARQL endpoints and query semantic data
- **Excel Integration**: Import, edit, and manage Excel datasheets with multi-sheet support
- **Tag Management**: Create and manage tags that associate nodes with datasheets
- **Data Mapping**: Map GraphDB properties to specific Excel cells
- **File Library**: Centralized datasheet management system

## Installation

### Prerequisites

- Python 3.8 or higher
- tkinter (usually included with Python)
- System packages for GUI support

### Setup

1. **Clone or download the application files**
   ```bash
   # Ensure you have the main application file: lean_digital_twin.py
   ```

2. **Install system dependencies (Ubuntu/Debian)**
   ```bash
   sudo apt update
   sudo apt install -y python3-venv python3-pip python3-tk
   ```

3. **Create and activate a virtual environment**
   ```bash
   python3 -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

4. **Install Python dependencies**
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Running the Application

```bash
# Activate virtual environment first
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Run the application
python lean_digital_twin.py
```

### Main Tabs

1. **Graphical Model**: Visualize relationships between GraphDB nodes
2. **GraphDB → Excel**: Connect to SPARQL endpoints and map data to Excel cells
3. **Datasheet Editor**: Edit Excel files with multi-sheet support
4. **Functionalities**: Manage tags and the datasheet library

### Basic Workflow

1. **Connect to GraphDB**:
   - Enter your SPARQL endpoint URL
   - Set the appropriate SPARQL prefix
   - Fetch all available nodes

2. **Create Property Mappings**:
   - Select a node and view its properties
   - Map properties to specific Excel cells
   - Add multiple property mappings as needed

3. **Select Target Datasheet**:
   - Import Excel files to the library
   - Choose a target datasheet and sheet name
   - Execute the data writing operation

4. **Manage Data**:
   - Use tags to organize nodes and datasheets
   - Edit datasheets directly in the application
   - Save changes back to the library

### File Management

- The application creates an `excel_files` directory for storing imported datasheets
- All file operations work within this managed library
- Files can be imported, edited, and removed through the GUI

## Dependencies

- **pandas**: Data manipulation and Excel file handling
- **pandastable**: Interactive table widget for Excel data
- **networkx**: Graph data structure and algorithms
- **matplotlib**: Graph visualization and plotting
- **SPARQLWrapper**: SPARQL query execution
- **openpyxl**: Excel file reading and writing

## Troubleshooting

### Common Issues

1. **Import Errors**: Ensure all dependencies are installed in the virtual environment
2. **GUI Issues**: Make sure tkinter is properly installed (python3-tk package)
3. **Excel Errors**: Verify that Excel files are valid and accessible
4. **SPARQL Errors**: Check network connectivity and endpoint URL validity

### Error Resolution

- **Missing modules**: Run `pip install -r requirements.txt` in your virtual environment
- **Permission errors**: Ensure write permissions in the application directory
- **Display issues**: For remote servers, ensure X11 forwarding is enabled

## License

This application is provided as-is for educational and research purposes.

## Support

For issues and questions, please check the error messages displayed in the application's status bar and message dialogs for detailed information about any problems.