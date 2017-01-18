package cn.yc;

import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.contenttype.ContentType;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.WordprocessingML.AlternativeFormatInputPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.org.apache.poi.util.IOUtils;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.CTAltChunk;

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

/**
 * Created by yuchao on 2017/1/11.
 */
public class MergeDocx {

    public static void main(String[] args) {
        List<InputStream> list = new ArrayList<InputStream>();
        try {
            list.add(new FileInputStream("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\0.docx"));
            list.add(new FileInputStream("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\1.docx"));
            list.add(new FileInputStream("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\2.docx"));
            list.add(new FileInputStream("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\3.docx"));
            list.add(new FileInputStream("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\4.docx"));
            list.add(new FileInputStream("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\5.docx"));
            list.add(new FileInputStream("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\6.docx"));
            list.add(new FileInputStream("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\7.docx"));
            list.add(new FileInputStream("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\8.docx"));
            list.add(new FileInputStream("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\9.docx"));
            list.add(new FileInputStream("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\10.docx"));
            list.add(new FileInputStream("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\11.docx"));
            list.add(new FileInputStream("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\12.docx"));
            list.add(new FileInputStream("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\13.docx"));
            list.add(new FileInputStream("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\14.docx"));
            list.add(new FileInputStream("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\15.docx"));
            list.add(new FileInputStream("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\16.docx"));
            list.add(new FileInputStream("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\17.docx"));
            list.add(new FileInputStream("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\18.docx"));
            list.add(new FileInputStream("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\19.docx"));
            list.add(new FileInputStream("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\20.docx"));
            list.add(new FileInputStream("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\21.docx"));
            list.add(new FileInputStream("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\22.docx"));
            list.add(new FileInputStream("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\23.docx"));
            list.add(new FileInputStream("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\24.docx"));
            list.add(new FileInputStream("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\25.docx"));
            list.add(new FileInputStream("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\26.docx"));
            list.add(new FileInputStream("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\27.docx"));
            InputStream inputStream = mergeDocx(list);

            FileOutputStream  fileOu = new FileOutputStream("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\merge.docx");
            byte[] b = new byte[1024];

            int hasRead = 0;
            while((hasRead = inputStream.read(b))>0)
            {
                fileOu.write(Arrays.copyOfRange(b,0,hasRead));

            }

            } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static InputStream mergeDocx(List<InputStream> streams)
            throws Docx4JException, IOException {

        WordprocessingMLPackage target = null;
        final File generated = File.createTempFile("generated", ".docx");

        int chunkId = 0;
        Iterator<InputStream> it = streams.iterator();
        while (it.hasNext()) {
            InputStream is = it.next();
            if (is != null) {
                if (target == null) {
                    // Copy first (master) document
                    OutputStream os = new FileOutputStream(generated);
                    os.write(IOUtils.toByteArray(is));
                    os.close();

                    target = WordprocessingMLPackage.load(generated);
                } else {
                    // Attach the others (Alternative input parts)
                    insertDocx(target.getMainDocumentPart(),
                            IOUtils.toByteArray(is), chunkId++);
                }
            }
        }

        if (target != null) {
            target.save(generated);
            return new FileInputStream(generated);
        } else {
            return null;
        }
    }

    // 插入文档
    private static void insertDocx(MainDocumentPart main, byte[] bytes, int chunkId) {
        try {
            AlternativeFormatInputPart afiPart = new AlternativeFormatInputPart(
                    new PartName("/part" + chunkId + ".docx"));
            afiPart.setContentType(new ContentType(""));
            afiPart.setBinaryData(bytes);
            Relationship altChunkRel = main.addTargetPart(afiPart);

            CTAltChunk chunk = Context.getWmlObjectFactory().createCTAltChunk();
            chunk.setId(altChunkRel.getId());

            main.addObject(chunk);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
