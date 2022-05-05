Azure Function - Generating MailMerged Word Document
-----------------------------------------------------
### Full article can be found on https://medium.com/@donnie.xyz.888/azure-function-generating-mailmerged-word-document-47b640e2c6fd

On this article, I will create an azure function that populate ms word document template mailmerge fields based on json data. This is quite common feature to generate an editable user friendly document, for example a draft of a contract that based on a standard template. This is a fully customizable  and open source alternative to Microsoft Power Automate Flow. Use case covered by this Function are: 1) uploading ms word template file; and 2) download generated ms word file based on the template and json data from http request body.
Azure Function - Generating MailMerged Word Document

---

# Exposed Azure Functions
There are two main features of this Azure Functions:
1. Upload word template
User can upload an msword document or template that contains mail merge field.
2. Download MailMerged document
By specifying the uploaded msWord file name and providing the mail merge field values in json format as the request body, the function will generate and returns the mail merged document.

[[MailMerge Azure Function Sequence Diagram]]

## Client Side
I will use postman to test the Function endpoint. On the real world scenario, the client usually 1) Web Browser or Mobile Front End , or 2) Other Function or Microservice.
## File Storage
On this article the uplopaded files are stored on the Function's temporary file, not on Azure Storage, because StoreFile* and ReadFile* process above is not implemented yet. This means the uploaded file will be removed when the function instance is restarted, and the uploaded file will only exists on the instance that received the upload http request. Lets improve this by using Azure Storage or other persistent mechanism on future article.

---

# Azure Function - Upload EndPoint
## Temp File
We will store the uploaded msword document into temporary file location. The temp file location must be guaranteed to be read-and-write-able by our upload function, and it should be transparent to OS which our upload function resides. To do this, we will use java's system property "java.io.tmpdir" as a base and then create a subdirectory as location for temporary file.
```
String BASEPATH = System.getProperty("java.io.tmpdir") + "ms-donniexyz-poi-azure-function-templates/";
```
## Upload HTTP Request Handling
As the Function is triggered by HTTP, we create a @HttpTrigger function endpoint with dataType = "binary".
To properly receive the msWord file, the request body must have content-type="multipart/form-data; boundary=xxx". From the boundary we can get the file size in bytes. I will use commons-fileupload MultipartStream to parse and stream request body input.
```
if (!"multipart/form-data".equals(contentType.split(";")[0])) {
 return
   request.createResponseBuilder(HttpStatus.BAD_REQUEST).build();
}
byte[] bs = request.getBody().get();
// Convert body to an input stream
InputStream in = new ByteArrayInputStream(bs); 
// Get boundary from content-type header
String boundary = contentType.split(";")[1].split("=")[1];
int bufSize = 1024;
MultipartStream multipartStream = new MultipartStream(in, boundary.getBytes(), bufSize, null);
```
## Postman Upload Request
To add postman file upload request, use request Body type "form-data", and add a field type = "File" and then select file to be uploaded.
[[Postman - form data for file upload]]

---

# Azure Function - MailMerge End Point
## MsWord Document
msword document file format to be processed is .docx or .dotx, that contains mail merge fields. To add a mail merge field into an msword document click "Insert" -> Quick Parts -> Field… -> Categories: Mail Merge -> Field Names: MergeField -> Field Name -> Input field name.
Microsoft word - add mail merge fieldThe field name is the json key of the request body, and period "." can be use to access child node. For example you can use mailmerge field name Contract.Amount to get the value from json"Contract": { "Amount": "123.00" }. 
Structured mailmerge field
## POI Lib
I will use apache poi library to read and write msword document file, specifically XWPF API.
To perform 'mail merging' using POI, first we need to know how to get the location for mailmerge fields. There are two variations of how mailmerge field stored inside the msword docx xml structure.
1. fldSimple variation

On this variation, the mail merge field definition and text value are located under fldSimple structure.
XPath example: /pkg:package/pkg:part/pkg:xmlData/w:document/w:body/w:p/w:fldSimple
2. 'three leaves' variation
On this variation, the merge field definition and the text value are located under different structure.
MailMerge field definition XPath example: /pkg:package/pkg:part/pkg:xmlData/w:document/w:body/w:tbl/w:tr/w:tc/w:p/w:r/w:instrText
MailMerge field text value XPath example: /pkg:package/pkg:part/pkg:xmlData/w:document/w:body/w:tbl/w:tr/w:tc/w:p/w:r/w:t, with a separator inbetween.

### Paragraph and Table
Notice the variations above are located under a "paragraph body" which we can use POI API XWPFDocument.getParagraphs()to retrieve it. Besides "paragraph body", the mailmerge field can be located inside a table, which can be retrieved using XWPFDocument.getTables().We also can use XWPFDocument.getBodyElements()to retrieve both paragraph and table and then use bodyElement.getElementType()to differentiate.
The process to get all of the mailmerge field is on our code PoiMailMerge.perform(), with main logic as follow:
To replace the mailmerge string value we can use POI APIXmlAnySimpleType.setStringValue().
## JSON MailMerge FieldValues
The json request body is simply converted into a map of map structure, and those values will replace the mailMerge fields. Mailmerge field that is not exist on request body json will be silently treated as blank "".
[[Mailmerge request body / json data exampleResponse]]
For azure function to return the binary msword file, convert the generated file into byte array, and then set it as the response body. 
To tell the client the download file name, set Content-Disposition=inline; filename=[XXXfileName] on the response header.
[[Generate the HTTP response]]

---

# Local Testing
The source code has been uploaded into public github repository ms-donniexyz-poi-azure-function, which does not contains any deployment and credential information. You can clone and run this in your local development environment. 
There is a ready to use postman script to speed up development and testing.
Ready to use postman scriptPostman upload request and response examplePostman generate mailmerged document exampleExample of msWord template document with mailmerge fields; ms-donniexyz-poi-azure-function\docs\template\PrivateHousingLoan_OfferLetter.docxExample of mailmerge json data. Notice the json key and structure is freely defined following each template document.
```
{
  "RefId": "REF_00001",
  "PrintDate": "Saturday, 3rd May 2022",
  "AddressName": "Mr Jay",
  "AddressLine1": "33X-6-3 The Garden",
  "AddressLine2": "Apartement Complex",
  "AddressCity": "Kuala Lumpur",
  "AddressPostCode": "55555",
  "Contract": {
    "Currency": "USD",
    "Amount": 123000,
    "Applicant": {
      "Name": "Mrs Jay"
    },
    "PropertyPrice": "250000",
    "Fee1Amount": "150",
    "DiscountAmount": "2500",
    "DurationInMonth": 120,
    "Installment": {
      "StartDate": "2022-08-01",
      "Amount": "1230",
      "InstallmentAmount": "1230"
    }
  },
  "ManagerName": "Joy Royce",
  "ManagerPosition": "Head of Staff",
  "Organization": "2MDB"
}
```
[[Example of generated msWord document; ms-donniexyz-poi-azure-function\docs\template\PrivateHousingLoan_OfferLetter_res.docx]]

---

# Azure Deployment
Clone the public github repository above into a private Azure Devops Project Repository and setup your deployment and credential information. I will create a separate article for this Azure Devops Deployment. Stay tuned.

---

# Conclusion
On this article it is shown how to create an a mailmerge msWord document generator using highly programmable and scalable azure function. The input for the process is an msWord document and json data, triggered via HTTP POST, which is fairly easy to be integrated with other microservice or functions. 
For source code, refer to github ms-donniexyz-poi-azure-function.
For postman script , refer to az Poi function.postman_collection.json.
For ms word document example, refer to PrivateHousingLoan_OfferLetter.docx.
