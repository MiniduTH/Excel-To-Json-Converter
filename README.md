# Excel to JSON Converter

A Java desktop application that converts Excel files to JSON format with a modern, user-friendly interface. The application allows users to select specific columns from Excel sheets and convert them into structured JSON data.

![Application Screenshot](https://i.ibb.co/cLBCmz1/Convertor-before-conversion.png)
![Application Screenshot](https://i.ibb.co/FYmr4R8/Convertor-after-conversion.png)
![Application Screenshot](https://i.ibb.co/JkhjvFv/converted-file.png)

## Features

- 📊 Convert Excel files (.xls, .xlsx) to JSON format
- 🎯 Select specific columns for conversion
- 📑 Support for custom sheet names
- 💫 Modern and intuitive user interface
- 💻 Cross-platform compatibility

## Prerequisites

- Java Runtime Environment (JRE) 11 or higher
- For building from source:
  - JDK 11 or higher
  - Apache Maven

## Usage

1. Launch the application
2. Click "Browse" to select your Excel file
3. Enter the sheet name (defaults to "Sheet1" if left empty)
4. Enter column numbers to convert (comma-separated)
5. Select output location
6. Click "Convert to JSON" to process the file

### Example

For an Excel file with columns:
```
| Name | Age | City |
|------|-----|------|
| John | 25  | NY   |
```

Enter "0,1,2" in the columns field to convert all three columns.

## Technical Details

The project consists of two main components:

1. **Core Logic** (Developed by me):
   - Excel file reading using Apache POI
   - Data extraction and processing
   - JSON conversion and file writing
   - Error handling and validation

2. **User Interface** (Designed by Claude AI):
   - Modern Swing-based UI
   - Intuitive layout and controls
   - Real-time status updates
   - Error feedback

### Dependencies

- Apache POI: Excel file handling
- JSON Library: JSON processing
- Swing: User interface
- JUnit: Testing

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- UI Design: [Claude AI (Anthropic)](https://www.anthropic.com)
- Excel Processing: [Apache POI](https://poi.apache.org)

## Contact

Your Name - [Your LinkedIn](https://linkedin.com/in/minidu0th)

Project Link: [https://github.com/MiniduTH/Excel-To-Json-Converter](https://github.com/MiniduTH/Excel-To-Json-Converter)
