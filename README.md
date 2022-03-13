# Java DOCX manipulator

### All libs used are in the lib folder and will have their own licences and stuff so do some looking before throwing any of this into something important
- slf4j (api, simple) https://www.slf4j.org/
- Apache commons (codec, logging) https://commons.apache.org/
- httpcore, httpclient https://hc.apache.org/index.html
- JODConverter  (https://github.com/sbraconnier/jodconverter)
- GSON (https://github.com/google/gson)
- Openoffice uno (http://www.openoffice.org/udk/)
- libreoffice unoil (https://www.libreoffice.org/)

### Features
replace placeholders `[placeholder]` with information using a Map of attributes and generate an updated docx from a template 

```
Map<String, String> attributes = new HashMap<String, String>();     
attributes.put("number", "22");                                           
attributes.put("text", "Something New Here");                     
String updatedDoc = replacePlaceholders("template.docx", attributes);
```

Convert docx to pdf using JODConverter library (https://github.com/sbraconnier/jodconverter)
```
String pdf = convertToPDF("document.docx");
```
