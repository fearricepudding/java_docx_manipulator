# Java DOCX manipulator

### All libs used are in the lib folder and will have their own licences and stuff so do some looking before throwing any of this into something important
- slf4j (api, simple) https://www.slf4j.org/
- Apache commons (codec, logging) https://commons.apache.org/
- httpcore, httpclient https://hc.apache.org/index.html
- JODConverter  (https://github.com/sbraconnier/jodconverter)
- GSON (https://github.com/google/gson)
- Openoffice uno (http://www.openoffice.org/udk/)
- libreoffice unoil (https://www.libreoffice.org/)

### full example usage
```
import com.fearricepudding.Docman;
import java.util.*;

public class example{
	public static void main(String[] args){
		Map<String, String> attributes = new HashMap<String, String>();
		attributes.put("number", "22");
		attributes.put("text", "Example text");
		
		try{
			String updateDoc = Docman.replacePlaceholders("template.docx", attributes);
			System.out.println("Created document from template: "+updateDoc);
		}catch(Exception e){
			System.out.println("An error create document from tempalte: "+ e);
		}

		try{
			String pdf = Docman.convertToPDF("updated.docx");
			System.out.println("Created PDF from docx: "+pdf);
		}catch(Exception e){
			System.out.println("An error converting to PDF: "+e);
		}
	}
}
```

### Features
replace placeholders `[placeholder]` with information using a Map of attributes and generate an updated docx from a template 

```
Map<String, String> attributes = new HashMap<String, String>();
attributes.put("number", "22");
attributes.put("text", "Example text");

try{
  String updateDoc = Docman.replacePlaceholders("template.docx", attributes);
  System.out.println("Created document from template: "+updateDoc);
}catch(Exception e){
  System.out.println("An error create document from tempalte: "+ e);
}

```

Convert docx to pdf using JODConverter library (https://github.com/sbraconnier/jodconverter)
```
try{
  String pdf = Docman.convertToPDF("updated.docx");
  System.out.println("Created PDF from docx: "+pdf);
}catch(Exception e){
  System.out.println("An error converting to PDF: "+e);
}
```

### Running example
```
cd example
./run.sh
```
