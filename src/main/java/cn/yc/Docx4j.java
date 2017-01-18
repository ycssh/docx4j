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
import java.util.Iterator;
import java.util.List;

/**
 * Created by yuchao on 2017/1/10.
 */
public class Docx4j {

    public void splitDocx(File f){

    }

    public InputStream mergeDocx(final List<InputStream> streams)
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
    private void insertDocx(MainDocumentPart main, byte[] bytes, int chunkId) {
//        try {
//            AlternativeFormatInputPart afiPart = new AlternativeFormatInputPart(
//                    new PartName("/part" + chunkId + ".docx"));
//            afiPart.setContentType(new ContentType(CONTENT_TYPE));
//            afiPart.setBinaryData(bytes);
//            Relationship altChunkRel = main.addTargetPart(afiPart);
//
//            CTAltChunk chunk = Context.getWmlObjectFactory().createCTAltChunk();
//            chunk.setId(altChunkRel.getId());
//
//            main.addObject(chunk);
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
    }
}
