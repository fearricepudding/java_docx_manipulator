# Java DOCX manipulator
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
