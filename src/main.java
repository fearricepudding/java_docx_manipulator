import java.util.zip.*;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.nio.file.Path;
import java.util.Random;
import org.jodconverter.local.*;
import org.jodconverter.core.office.*;
import org.jodconverter.local.office.*;
import org.jodconverter.remote.office.*;
import java.util.*;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import com.sun.star.lib.uno.helper.*;
import java.lang.IllegalArgumentException;
import com.google.gson.*;
import com.sun.star.comp.loader.JavaLoader;

import java.nio.file.Paths;

class docman{
	public static void main(String[] args) throws IOException{
		System.out.println("DOCMAN");

		Map<String, String> attributes = new HashMap<String, String>();
		attributes.put("number", "22");
		attributes.put("text", "Something New Here");

		String updatedDoc = replacePlaceholders("template.docx", attributes);
		String pdf = convertToPDF(updatedDoc);
	}

	public static String convertToPDF(String file){
		File inputFile = new File(file);
		String newName = file.replace(".docx", ".pdf");
		File outputFile = new File(newName);

		//final RemoteOfficeManager officeManager = RemoteOfficeManager.install("127.0.0.1");
		final LocalOfficeManager officeManager = LocalOfficeManager
													//.officeHome("")
													//.build();
													.install();
		try {
			officeManager.start();

			JodConverter
					 .convert(inputFile)
					 .to(outputFile)
					 .execute();
		} catch(OfficeException e){
			e.printStackTrace();
		} finally {
			OfficeUtils.stopQuietly(officeManager);
		}
		return newName;
	}

	public static String replacePlaceholders(String template, Map<String, String> data) throws IOException{
		String tempLocation = unzip(template);
		String contentPath = tempLocation+"/word/document.xml";
		String content = getFileContent(contentPath);

		for(Map.Entry<String, String> entry : data.entrySet()){
			String selector = "\\["+entry.getKey()+"\\]";
			String value = entry.getValue();
			content = content.replaceAll(selector, value);
		}

		updateFileContent(contentPath, content);
		Path toZip = Paths.get(tempLocation);
		String newDocName = "updated.docx"; // replace with random name
		zipFolder(toZip, newDocName);
		deleteDir(tempLocation);
		return newDocName;
	}

	public static void updateFileContent(String filePath, String fileContent) throws IOException{
		Path file = Paths.get(filePath);
		Files.write(file, fileContent.getBytes());
	}

	public static String getFileContent(String filePath){
		String content = "";
		try{
			content = new String ( Files.readAllBytes( Paths.get(filePath) ) );
		}catch (IOException e){
			e.printStackTrace();
		}
		return content;
	}

	public static void zipFolder(Path source, String zipFileName) throws IOException {
        try(ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(zipFileName))) {
            Files.walkFileTree(source, new SimpleFileVisitor<Path>() {
                @Override
                public FileVisitResult visitFile(Path file,
                    BasicFileAttributes attributes) {
                    if (attributes.isSymbolicLink()) {
                        return FileVisitResult.CONTINUE;
                    }
                    try (FileInputStream fis = new FileInputStream(file.toFile())) {
                        Path targetFile = source.relativize(file);
                        zos.putNextEntry(new ZipEntry(targetFile.toString()));
                        byte[] buffer = new byte[1024];
                        int len;
                        while ((len = fis.read(buffer)) > 0) {
                            zos.write(buffer, 0, len);
                        }
                        zos.closeEntry();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                    return FileVisitResult.CONTINUE;
                }
                @Override
                public FileVisitResult visitFileFailed(Path file, IOException exc) {
                    return FileVisitResult.CONTINUE;
                }
            });
        }
    }

	public static void deleteDir(String path){
		File source = new File(path);
		File[] content = source.listFiles();
		if(content != null){
			for(File f : content){
				if(!Files.isSymbolicLink(f.toPath())){
					deleteDir(f.toString());
				}
			}
		}
		source.delete();
	}

	public static String randomDir(){
		byte[] array = new byte[7]; // length is bounded by 7
		new Random().nextBytes(array);
		String generatedString = new String(array, Charset.forName("UTF-8"));
		return generatedString;
	}

	public static String unzip(String fileZip) throws IOException{
		String dest = "./tmp/"+randomDir();
        File destDir = new File(dest);
        byte[] buffer = new byte[1024];
        ZipInputStream zis = new ZipInputStream(new FileInputStream(fileZip));
        ZipEntry zipEntry = zis.getNextEntry();
        while (zipEntry != null) {
			File newFile = newFile(destDir, zipEntry);
			 if (zipEntry.isDirectory()) {
				 if (!newFile.isDirectory() && !newFile.mkdirs()) {
					 throw new IOException("Failed to create directory " + newFile);
				 }
			 } else {
				 File parent = newFile.getParentFile();
				 if (!parent.isDirectory() && !parent.mkdirs()) {
					 throw new IOException("Failed to create directory " + parent);
				 }
				 FileOutputStream fos = new FileOutputStream(newFile);
				 int len;
				 while ((len = zis.read(buffer)) > 0) {
					 fos.write(buffer, 0, len);
				 }
				 fos.close();
			 }
			 zipEntry = zis.getNextEntry();
        }
        zis.closeEntry();
        zis.close();
		return dest;
	};

	public static File newFile(File destinationDir, ZipEntry zipEntry) throws IOException {
		File destFile = new File(destinationDir, zipEntry.getName());
		String destDirPath = destinationDir.getCanonicalPath();
		String destFilePath = destFile.getCanonicalPath();
		if (!destFilePath.startsWith(destDirPath + File.separator)) {
			throw new IOException("Entry is outside of the target dir: " + zipEntry.getName());
		}
		return destFile;
	}
}
