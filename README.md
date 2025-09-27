# CREDIFLEXI Report Automation System

## Overview

CREDIFLEXI is a professional web application designed to automate credit portfolio antiquity reports processing. The system provides automated data processing, fraud filtering, geolocation integration, and professional Excel report generation with advanced formatting and conditional styling.

## Features

### Core Functionality
- **Automated Report Processing**: Processes Excel files containing credit portfolio antiquity data
- **Fraud Detection**: Automatic filtering of fraudulent account codes using predefined fraud lists
- **Geolocation Integration**: Generates Google Maps links for address verification
- **PAR Calculation**: Automatic calculation of Portfolio at Risk (PAR) percentages based on delinquency days
- **Multi-Sheet Reports**: Generates comprehensive Excel reports with multiple organized sheets

### Report Generation
- **Complete Report Sheet**: Full dataset with all processed records
- **Delinquency Sheet**: Filtered records showing accounts with 1+ days of delinquency
- **Coordination Sheets**: Individual sheets organized by coordination/region
- **Professional Formatting**: Conditional formatting, tables, and hyperlinks

### Data Processing Capabilities
- **Data Cleaning**: Standardizes phone numbers and data formats
- **Column Mapping**: Intelligent column detection and mapping
- **Duplicate Prevention**: Automatic handling of duplicate columns
- **Data Integrity**: Comprehensive data validation and integrity checks

## Technical Architecture

### Technology Stack
- **Backend**: Python 3.13 with Flask framework
- **Data Processing**: Pandas for data manipulation and analysis
- **Excel Generation**: OpenPyXL for advanced Excel file creation
- **Frontend**: HTML5, CSS3, JavaScript with Bootstrap 5
- **Configuration**: Centralized configuration management

### Project Structure
```
automatizador-crediflexi/
├── app/
│   ├── __init__.py
│   └── reportes.py          # Core report processing logic
├── static/
│   ├── css/
│   │   └── style.css        # Application styling
│   ├── js/                  # JavaScript files
│   └── downloads/           # Generated report storage
├── templates/
│   ├── base.html           # Base template
│   ├── index.html          # Home page
│   └── antiguedad.html     # Report processing page
├── uploads/                # Input file storage
├── app.py                  # Flask application entry point
├── config.py              # Configuration settings
├── pipeline.py            # Data processing pipeline
└── README.md              # This file
```

## Installation

### Prerequisites
- Python 3.13 or higher
- pip package manager

### Setup Instructions

1. **Clone the repository**
   ```bash
   git clone <repository-url>
   cd automatizador-crediflexi
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Configure the application**
   - Review and modify `config.py` for your specific requirements
   - Update fraud list, column mappings, and other configurations as needed

4. **Run the application**
   ```bash
   python app.py
   ```

5. **Access the application**
   - Open your web browser and navigate to `http://localhost:5000`

## Configuration

### Key Configuration Files

#### config.py
Contains all application settings including:
- **LISTA_FRAUDE**: Fraudulent account codes for automatic filtering
- **COLUMN_MAPPING**: Column name mappings for different data formats
- **DTYPE_CONFIG**: Data type configurations for optimal processing
- **EXCEL_CONFIG**: Excel formatting and styling settings
- **Colors and Styling**: Professional color schemes and formatting rules

### Customization Options
- Fraud list management
- Column mapping adjustments
- Excel formatting preferences
- Color schemes and branding
- Processing parameters

## Usage

### Basic Workflow

1. **Access the Application**
   - Navigate to the home page
   - Click "Procesar" to access the report processing interface

2. **Upload Excel File**
   - Select your credit portfolio antiquity Excel file
   - Supported formats: .xlsx, .xls
   - Maximum file size: 16MB

3. **Automatic Processing**
   - The system automatically processes the uploaded file
   - Performs fraud filtering
   - Generates geolocation links
   - Calculates PAR percentages
   - Applies professional formatting

4. **Download Results**
   - Processed report is automatically generated
   - Download the Excel file with multiple organized sheets
   - File is saved with current date in filename

### Advanced Features

#### Fraud Filtering
- Automatic detection and removal of fraudulent account codes
- Configurable fraud list in `config.py`
- Comprehensive logging of filtered records

#### Geolocation Integration
- Automatic generation of Google Maps links
- Address validation and verification
- Clickable hyperlinks in Excel reports

#### PAR Calculation
- Automatic calculation based on delinquency days
- Intelligent column positioning
- Professional formatting with color coding

## File Formats

### Input Requirements
- **Format**: Excel files (.xlsx, .xls)
- **Required Columns**: Account code, delinquency days, coordination, address
- **Data Quality**: Clean, standardized data for optimal processing

### Output Specifications
- **Format**: Excel files (.xlsx)
- **Sheets**: Complete report, delinquency report, coordination-specific sheets
- **Features**: Conditional formatting, tables, hyperlinks, professional styling

## Error Handling

### Comprehensive Error Management
- File validation and format checking
- Data integrity verification
- Processing error recovery
- User-friendly error messages
- Detailed logging for debugging

### Common Issues and Solutions
- **File Format Issues**: Ensure Excel files are properly formatted
- **Missing Columns**: Verify required columns are present in input data
- **Data Quality**: Clean and standardize input data before processing
- **Memory Issues**: Process smaller files or increase system memory

## Security Features

### Data Protection
- Secure file upload handling
- Temporary file cleanup
- No persistent storage of sensitive data
- Configurable fraud filtering

### Input Validation
- File type verification
- Size limit enforcement
- Data format validation
- Malicious content detection

## Performance Optimization

### Processing Efficiency
- Optimized pandas operations
- Memory-efficient data handling
- Parallel processing capabilities
- Intelligent caching mechanisms

### Scalability Considerations
- Modular architecture for easy expansion
- Configurable processing parameters
- Resource usage monitoring
- Batch processing capabilities

## Maintenance and Support

### Logging and Monitoring
- Comprehensive logging system
- Processing statistics tracking
- Error monitoring and reporting
- Performance metrics collection

### Regular Maintenance
- Configuration updates
- Fraud list maintenance
- Performance optimization
- Security updates

## Contributing

### Development Guidelines
- Follow Python PEP 8 style guidelines
- Maintain comprehensive documentation
- Include unit tests for new features
- Ensure backward compatibility

### Code Quality Standards
- Clean, readable code structure
- Comprehensive error handling
- Detailed logging and debugging
- Professional documentation

## License

This project is proprietary software developed for CREDIFLEXI AS CV. All rights reserved.

## Contact Information

For technical support or inquiries regarding this application, please contact the development team.

---

**CREDIFLEXI** - Professional Credit Portfolio Management Solutions
