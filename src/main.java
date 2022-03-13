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

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import com.sun.star.lib.uno.helper.*;
import java.lang.IllegalArgumentException;
import com.google.gson.*;
import com.sun.star.comp.loader.JavaLoader;

class docman{
	public static void main(String[] args) throws IOException{
		System.out.println("DOCMAN");

		String temp = unzip("orig.docx");
		String contentPath = temp+"/word/document.xml";

		Path file = Path.of(contentPath);

		String content = Files.readString(file);

		System.out.println(content.contains("[number]"));
		String newContent = content.replaceAll("\\[number\\]", "22");
		Files.write(file, newContent.getBytes());

		Path toZip = Path.of(temp);
	    zipFolder(toZip, "test.docx");

		File inputFile = new File("test.docx");
		File outputFile = new File("document.pdf");

		//final RemoteOfficeManager officeManager = RemoteOfficeManager.install("127.0.0.1");
		final LocalOfficeManager officeManager = LocalOfficeManager.install();
		try {
			officeManager.start();

			JodConverter
					 .convert(inputFile)
					 .to(outputFile)
					 .execute();
		} catch(OfficeException e){
			//
		} finally {
			OfficeUtils.stopQuietly(officeManager);
		}
	}

	public static void zipFolder(Path source, String zipFileName) throws IOException {
        try(ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(zipFileName))) {
            Files.walkFileTree(source, new SimpleFileVisitor<>() {
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
                    System.err.printf("Unable to zip : %s%n%s%n", file, exc);
                    return FileVisitResult.CONTINUE;
                }
            });
        }
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
