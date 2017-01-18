package cn.yc;

import org.docx4j.Docx4J;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.Parts;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPart;
import org.docx4j.openpackaging.parts.WordprocessingML.DocumentSettingsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.relationships.Relationships;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by yuchao on 2017/1/11.
 */
public class Docx2html {

    public static void docxToHtml(String filepath, String outpath) throws Docx4JException, FileNotFoundException {
        WordprocessingMLPackage wmp = WordprocessingMLPackage.load(new File(filepath));
        Docx4J.toHTML(wmp, "html/resources", "resources", new FileOutputStream(new File(outpath)));
    }

    public static void main(String[] args) {
        try {
            WordprocessingMLPackage wordMLPackage = Docx4J.load(new File("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\27.docx"));


            MainDocumentPart documentPart1 = wordMLPackage.getMainDocumentPart();

            RelationshipsPart relsPart1 = documentPart1.getRelationshipsPart();
            Relationships rels1 = relsPart1.getRelationships();
            List<Relationship> relsList1 = rels1.getRelationship();
            for(Relationship relationship:relsList1){
                System.out.println(relationship.getType()+"\t"+relationship.getTarget()+"\t"+relationship.getTargetMode());
                System.out.println(relationship.getParent());
            }


//            HashMap<PartName, Part> parts = wordMLPackage.getParts().getParts();
//
//            RelationshipsPart relsPart = wordMLPackage.getMainDocumentPart().getRelationshipsPart();
//            Relationships rels = relsPart.getRelationships();
//            List<Relationship> relsList = rels.getRelationship();
//            for(Map.Entry<PartName, Part> entry:parts.entrySet()){
//                PartName partName = entry.getKey();
//                Part part = entry.getValue();
//                BinaryPart oPart = (BinaryPart) wordMLPackage.getParts().getParts().get(partName);
//                System.out.println(new String(oPart.getBytes()));
//                oPart.getRelationshipsPart();
//                System.out.println(partName.getName());
//            }

//            getPart.save(new File("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\28.docx"));

//            docxToHtml("C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\a.docx","C:\\Users\\yuchao\\Desktop\\word\\aaa\\a\\html.html");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
