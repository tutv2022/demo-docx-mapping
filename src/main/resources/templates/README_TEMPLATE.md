# Template Directory

Place your DOCX template files in this directory.

## Template Format

Your DOCX template should contain placeholders in the following format:

```
${variableName}
```

## Example Template Content

Create a DOCX file (e.g., `template.docx`) with content like:

```
Dear ${customerName},

Your order ${orderNumber} has been processed on ${orderDate}.

Order Details:
- Item: ${itemName}
- Quantity: ${quantity}
- Total Amount: ${totalAmount}

Thank you for your business!

Best regards,
${companyName}
```

## Usage

Once you have created your template file, you can reference it in the API call:

```bash
curl -X POST "http://localhost:8080/api/documents/process?template=template.docx" \
  -H "Content-Type: application/json" \
  -d '{
    "variables": {
      "customerName": "John Doe",
      "orderNumber": "ORD-12345",
      "orderDate": "2024-01-15",
      "itemName": "Product XYZ",
      "quantity": "2",
      "totalAmount": "$199.99",
      "companyName": "Your Company"
    }
  }' \
  --output result.docx
```

