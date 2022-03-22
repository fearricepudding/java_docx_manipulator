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
