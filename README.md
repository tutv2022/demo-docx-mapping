# DOCX Template Processor - Spring Boot POC

A Proof of Concept project demonstrating how to process DOCX template files using Apache POI in a Spring Boot application. This project replaces template variables (placeholders) with actual data while preserving the original document formatting.

## Features

- ✅ Process DOCX templates with placeholder variables (e.g., `${variableName}`)
- ✅ Preserve original document formatting (bold, italic, colors, fonts, etc.)
- ✅ Process all document sections (body, headers, footers, tables)
- ✅ REST API for template processing
- ✅ Support for both classpath templates and file uploads

## Prerequisites

- Java 17 or higher
- Maven 3.6 or higher

## Project Structure

```
demo-docx-apache-POI/
├── pom.xml
├── src/
│   └── main/
│       ├── java/
│       │   └── com/
│       │       └── example/
│       │           └── docxprocessor/
│       │               ├── DocxProcessorApplication.java
│       │               ├── controller/
│       │               │   └── DocumentController.java
│       │               ├── service/
│       │               │   └── DocxTemplateService.java
│       │               └── model/
│       │                   └── TemplateData.java
│       └── resources/
│           ├── application.properties
│           └── templates/
│               └── (your template.docx file goes here)
```

## Setup

1. **Clone or navigate to the project directory**

2. **Build the project**
   ```bash
   mvn clean install
   ```

3. **Create a template DOCX file**
   - Create a DOCX file with placeholders in the format `${variableName}`
   - Example content: "Hello ${name}, your order ${orderId} is ready."
   - Place it in `src/main/resources/templates/template.docx`

4. **Run the application**
   ```bash
   mvn spring-boot:run
   ```

   The application will start on `http://localhost:8080`

## Usage

### API Endpoints

#### 1. Health Check
```bash
GET http://localhost:8080/api/documents/health
```

#### 2. Process Template (from classpath)

**Using curl:**
```bash
curl -X POST "http://localhost:8080/api/documents/process?template=template.docx" \
  -H "Content-Type: application/json" \
  -d '{
    "variables": {
      "name": "John Doe",
      "orderId": "12345",
      "date": "2024-01-15"
    }
  }' \
  --output result.docx
```

**Using Postman:**
- Method: POST
- URL: `http://localhost:8080/api/documents/process?template=template.docx`
- Headers: `Content-Type: application/json`
- Body (JSON):
  ```json
  {
    "variables": {
      "name": "John Doe",
      "orderId": "12345",
      "date": "2024-01-15"
    }
  }
  ```

#### 3. Process Template (upload file)

**Using curl:**
```bash
curl -X POST "http://localhost:8080/api/documents/process" \
  -F "file=@/path/to/your/template.docx" \
  -F 'data={"variables": {"name": "John Doe", "orderId": "12345"}}' \
  --output result.docx
```

**Note:** The file upload approach requires multipart/form-data format. For JSON body with file upload, you may need to use a different approach or modify the controller.

## Template Format

Placeholders in your DOCX template should follow this format:
- `${variableName}` - Will be replaced with the value from the variables map

**Example template content:**
```
Dear ${customerName},

Your order ${orderNumber} has been processed.

Order Details:
- Item: ${itemName}
- Quantity: ${quantity}
- Total: ${totalAmount}

Thank you for your business!

Best regards,
${companyName}
```

## How It Works

1. **Template Reading**: The service reads the DOCX file using Apache POI's `XWPFDocument`
2. **Placeholder Detection**: Uses regex pattern `\$\{([^}]+)\}` to find placeholders
3. **Replacement**: Replaces placeholders with actual values from the provided data map
4. **Format Preservation**: 
   - Preserves formatting by copying run properties (bold, italic, font, color, etc.)
   - Maintains paragraph structure
   - Processes all document sections (body, headers, footers, tables)
5. **Output**: Returns the processed document as a byte array

## Technical Details

- **Spring Boot**: 3.2.0
- **Apache POI**: 5.2.5 (for DOCX processing)
- **Java**: 17
- **Maven**: For dependency management

## Limitations

- Placeholders must be in the format `${variableName}` (case-sensitive)
- Complex formatting across multiple runs may not be perfectly preserved
- Very large documents (>10MB) may need configuration adjustments
- Nested placeholders or complex expressions are not supported

## Future Enhancements

- Support for nested objects in template data
- Support for loops/repeating sections
- Support for conditional sections
- Better formatting preservation for complex documents
- Template validation endpoint
- Batch processing support

## License

This is a POC project for demonstration purposes.

