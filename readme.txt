import java.io.*;
import java.nio.charset.Charset;
import java.util.Enumeration;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

public class DocxHelper {

    public static void saveXmlToDocx(File sourceDocxFile, String targetDocxFilePath, String newXmlString) throws IOException {

        ZipFile original = new ZipFile(sourceDocxFile);
        ZipOutputStream outputStream = new ZipOutputStream(new FileOutputStream(targetDocxFilePath));

        Enumeration entries = original.entries();
        while (entries.hasMoreElements()) {

            ZipEntry entry = (ZipEntry)entries.nextElement();
            ZipEntry newEntry = new ZipEntry(entry.getName());

            outputStream.putNextEntry(newEntry);
            if  ("word/document.xml".equalsIgnoreCase(entry.getName())) {

                byte[] bufferOut = new byte[512];
                InputStream in = new ByteArrayInputStream(newXmlString.getBytes(Charset.forName("UTF-8")));
                while (0 < in.available()){
                    int read = in.read(bufferOut);
                    outputStream.write(bufferOut,0,read);
                }
                in.close();
            }
            else{

                byte[] buffer = new byte[(int) entry.getSize()];
                InputStream in = original.getInputStream(entry);
                while (0 < in.available()){
                    int read = in.read(buffer);
                    outputStream.write(buffer,0,read);
                }
                in.close();
            }
            outputStream.closeEntry();
        }

        outputStream.close();
        original.close();
    }

    public static void saveXmlToDocx(String sourceDocxFilePath, String targetDocxFilePath, String newXmlString) throws IOException {

        saveXmlToDocx(new File(sourceDocxFilePath), targetDocxFilePath, newXmlString);
    }

    public static String loadXmlFromDocx(File sourceDocxFile) throws IOException {

        ZipFile zip = new ZipFile(sourceDocxFile);
        InputStream isXml = zip.getInputStream(zip.getEntry("word/document.xml"));
        BufferedReader br = new BufferedReader(new InputStreamReader(isXml, "UTF-8"));

        String part;
        StringBuilder out = new StringBuilder();
        while((part = br.readLine()) != null)
            out.append(part);

        br.close();
        isXml.close();
        zip.close();

        return out.toString();
    }

    public static String loadXmlFromDocx(String sourceDocxFilePath) throws IOException {

        return loadXmlFromDocx(new File(sourceDocxFilePath));
    }

}
