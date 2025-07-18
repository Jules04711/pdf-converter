# Universal Document to PDF Converter

A modern, secure, and user-friendly Streamlit application for converting multiple document formats (Word, PowerPoint, Text, Markdown) to PDF format.

## Features

- ✅ **Multiple Format Support**: Word (.docx), PowerPoint (.pptx), Text (.txt), Markdown (.md)
- ✅ **Secure File Validation**: Validates file format, size, and content
- ✅ **Real-time Conversion**: Fast conversion with progress tracking
- ✅ **User-friendly Interface**: Clean, intuitive design with helpful feedback
- ✅ **Conversion Analytics**: Track conversion time and file size changes
- ✅ **Session History**: View recent conversions
- ✅ **Automatic Cleanup**: Secure temporary file management
- ✅ **Error Handling**: Comprehensive error messages and troubleshooting tips

## System Requirements

### Software Requirements
- **Word Documents**: Microsoft Word 2007 or later (Windows/macOS)
- **PowerPoint Documents**: Python-pptx library (included in requirements)
- **Text Files**: ReportLab library (included in requirements)
- **Markdown Files**: Markdown + WeasyPrint libraries (included in requirements)

### Python Requirements
- Python 3.7 or higher
- See `requirements.txt` for package dependencies

## Installation

1. **Clone or download this repository**
   ```bash
   git clone <repository-url>
   cd pdf-converter
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Ensure Microsoft Word is installed and licensed**
   - The app requires Word to be installed on the system
   - Make sure Word can be launched and is properly licensed

## Usage

1. **Start the application**
   ```bash
   streamlit run app.py
   ```

2. **Open your browser**
   - The app will automatically open in your default browser
   - If not, navigate to `http://localhost:8501`

3. **Convert your documents**
   - Upload a document file (max 50MB)
   - Click "Convert to PDF"
   - Download your converted PDF file

## File Specifications

- **Input Formats**: 
  - `.docx` (Microsoft Word 2007+)
  - `.pptx` (Microsoft PowerPoint 2007+)
  - `.txt` (Plain text files)
  - `.md` (Markdown files)
- **Output Format**: `.pdf` (Portable Document Format)
- **Maximum File Size**: 50MB
- **Supported Content**: Text, images, tables, formatting (varies by input type)

## Security Features

- File format validation
- File size limits
- Content validation (checks for valid DOCX structure)
- Automatic temporary file cleanup
- Secure filename sanitization

## Troubleshooting

### Common Issues

1. **"Conversion failed" error**
   - Ensure Microsoft Word is installed and licensed
   - Check that the DOCX file isn't corrupted
   - Try with a different Word document

2. **"Invalid DOCX file format" error**
   - Make sure the file is a valid .docx file
   - Try opening the file in Word first to verify it's not corrupted

3. **"File too large" error**
   - Reduce file size or split into smaller documents
   - Maximum supported size is 50MB

### System-Specific Notes

- **Windows**: Works best with Microsoft Office installed
- **macOS**: Requires Microsoft Word for Mac
- **Linux**: Consider using LibreOffice-based alternatives

## Development

### Project Structure
```
pdf-converter/
├── app.py              # Main Streamlit application
├── requirements.txt    # Python dependencies
└── README.md          # This file
```

### Key Functions
- `validate_file()`: File validation and security checks
- `convert_word_to_pdf()`: Core conversion logic
- `get_download_link()`: Secure PDF download generation
- `cleanup_temp_files()`: Temporary file management

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## License

This project is open source. Please check the license file for details.

## Support

For issues and questions:
1. Check the troubleshooting section above
2. Ensure all system requirements are met
3. Create an issue in the repository

---

Built with ❤️ using Streamlit and docx2pdf 