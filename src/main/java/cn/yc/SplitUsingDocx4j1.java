package cn.yc;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.docx4j.XmlUtils;
import org.docx4j.dml.CTBlip;
import org.docx4j.openpackaging.contenttype.ContentType;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.org.apache.xml.serializer.utils.SystemIDResolver;
import org.docx4j.relationships.Relationship;
import org.docx4j.relationships.Relationships;
import org.docx4j.utils.SingleTraversalUtilVisitorCallback;
import org.docx4j.utils.TraversalUtilVisitor;
import org.docx4j.wml.ArrayListWml;
import org.docx4j.wml.Body;
import org.docx4j.wml.P;
import org.docx4j.wml.Styles;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.regex.Pattern;

public class SplitUsingDocx4j1 {

    /**
     * @param args
     * @throws Docx4JException
     * @throws FileNotFoundException
     */
    public static void main(String[] args) throws Docx4JException,
            IOException, JAXBException {
        File dir = new File("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\a.docx");
        FileWriter fw = new FileWriter("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\a.txt");
        // Creating new documents
        WordprocessingMLPackage doc1 = WordprocessingMLPackage.createPackage();

        // loading existing document
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage
                .load(new java.io.File(dir.getPath()));

        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
        List<Object> obj = wordMLPackage.getMainDocumentPart().getContent();

        RelationshipsPart relsPart = documentPart.getRelationshipsPart();
        Relationships rels = relsPart.getRelationships();
        List<Relationship> relsList = rels.getRelationship();
        StyleDefinitionsPart sdp = documentPart.getStyleDefinitionsPart();
        Styles tempStyle = sdp.getJaxbElement();
        doc1.getMainDocumentPart().getStyleDefinitionsPart()
                .setJaxbElement(tempStyle);

        List<Object> list = new ArrayList();
        List<Integer> split = new ArrayList<Integer>();
        for (int i = 0; i < obj.size(); i++) {
            if (isSmallTilete(obj.get(i).toString())) {
                split.add(i);
            }
            list.add(obj.get(i));
        }
        List<String> images = new ArrayList<String>();
        for (Relationship r : relsList) {
            if (r.getType().equals(Namespaces.IMAGE)
                    && (r.getTargetMode() == null
                    || r.getTargetMode().equalsIgnoreCase("internal"))) {
                images.add(r.getId());
                r.setTargetMode("External");
            }
        }

        int k = 0;
        for (int i = 0; i < split.size(); i++) {
            if (i < split.size() - 1) {
                for (int j = split.get(i); j < split.get(i + 1); j++) {
                    doc1.getMainDocumentPart().addObject(list.get(j));
                }
                filterImage(images,wordMLPackage,doc1);
                doc1.save(new File("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\" + k + ".docx"));

                doc1 = WordprocessingMLPackage.createPackage();
                doc1.getMainDocumentPart().getStyleDefinitionsPart().setJaxbElement(tempStyle);
                k++;
            } else {
                for (int j = split.get(i); j < list.size(); j++) {
                    doc1.getMainDocumentPart().addObject(list.get(j));
                }
                filterImage(images,wordMLPackage,doc1);
                doc1.save(new File("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\" + k + ".docx"));
            }
        }
        fw.flush();
    }

    public static boolean isSmallTilete(String str) {
        Pattern pattern = Pattern.compile("^([\\d]+[-\\.].*)");
        Pattern pattern1 = Pattern.compile("^([\\d]+[-\\:].*)");
        Pattern pattern3 = Pattern.compile("^([\\d]+[-\\：].*)");
        Pattern pattern2 = Pattern.compile("^([\\d]+[-\\、].*)");
        Pattern pattern4 = Pattern.compile("^([\\d]+[-\\．].*)");
        boolean result = pattern.matcher(str).matches() || pattern1.matcher(str).matches() ||
                pattern2.matcher(str).matches() || pattern3.matcher(str).matches() || pattern4.matcher(str).matches();
        return result;
    }

    //判断Str是否是大标题
    public static boolean isBigTilete(String str) {
        boolean iso = false;
        if (str.contains("一、")) {
            iso = true;
        } else if (str.contains("二、")) {
            iso = true;
        } else if (str.contains("三、")) {
            iso = true;
        } else if (str.contains("四、")) {
            iso = true;
        } else if (str.contains("五、")) {
            iso = true;
        } else if (str.contains("六、")) {
            iso = true;
        } else if (str.contains("七、")) {
            iso = true;
        } else if (str.contains("八、")) {
            iso = true;
        }
        return iso;
    }

    /**
     * <w:drawing>
     * <wp:inline distT="0" distB="0" distL="0" distR="0">
     * <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
     * <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
     * <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
     * <pic:blipFill>
     * <a:blip r:embed="rId5" />
     */
    public static class TraversalUtilBlipVisitor extends TraversalUtilVisitor<CTBlip> {

        @Override
        public void apply(CTBlip element, Object parent, List<Object> siblings) {

            if (element.getEmbed() != null) {

                String relId = element.getEmbed();
                // Add r:link
                element.setLink(relId);
                // Remove r:embed
                element.setEmbed(null);

            }
        }

    }

    private static void filterImage(List<String> images, WordprocessingMLPackage wordMLPackage, WordprocessingMLPackage doc1) throws InvalidFormatException, IOException {

//        FileWriter fw = new FileWriter("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\"+ System.currentTimeMillis()+".txt");
        String s = XmlUtils.marshaltoString(doc1.getMainDocumentPart().getJaxbElement(), true, true);
//        System.out.println(s);
//        fw.write(s);
//        fw.flush();
        for (String id : images) {
            if(s.contains(id)){
                Relationship re = wordMLPackage.getMainDocumentPart().getRelationshipsPart().getRelationshipByID(id);
                PartName partName = new PartName("/word/" + re.getTarget());
                BinaryPart oPart = (BinaryPart) wordMLPackage.getParts().getParts().get(new PartName("/word/" + re.getTarget()));
                BinaryPart bPart = new BinaryPart(partName);
                bPart.setBinaryData(oPart.getBytes());
                bPart.setContentType(new ContentType(oPart.getContentType()));
                bPart.setRelationshipType(re.getType()/* Namespaces.IMAGE */);
                Relationship newRe = doc1.getMainDocumentPart().addTargetPart(bPart);
                newRe.setId(id);
                newRe.setType(re.getType());
                System.out.println(id+"\t"+re.getType());
            }
        }
    }

}