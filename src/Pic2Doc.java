
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlToken;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualDrawingProps;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;

import java.io.IOException;
import java.io.InputStream;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.imageio.ImageIO;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class Pic2Doc extends XWPFDocument {

    static String destFile = "a2.docx";
    static String imagesPath = "C:\\pic2doc\\photos";
    static String programPath = "C:\\pic2doc";
    static String template = "templateFile.docx";

    public static void main(String[] args) throws FileNotFoundException, IOException, InvalidFormatException {
        createTemplateFile();
        createDestFile();
        CustomXWPFDocument document = new CustomXWPFDocument(new FileInputStream(new File(programPath + "\\Template\\" + template)));
        getAllFiles(imagesPath, document);

    }

    public static void addString(String s, String path, XWPFDocument document) throws FileNotFoundException, IOException {
        File f = new File(s);
        String name = f.getName();
        name = name.substring(0, name.length() - 4);
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText(name);
        run.setFontSize(33);
        FileOutputStream out = new FileOutputStream(path);
        document.write(out);
        out.close();
    }

    public static void getAllFiles(String file, CustomXWPFDocument document) throws IOException {
        File f = new File(file);

        if (f.isDirectory()) {
            System.out.println(f.getName());
            File[] listOfFiles = f.listFiles();
            for (int i = 0; i < listOfFiles.length; i++) {
                System.out.println(listOfFiles[i].getName());

                if (listOfFiles[i].isFile() && listOfFiles[i].getName().toLowerCase().contains("jpg")) {
                    addString(listOfFiles[i].getAbsolutePath(), programPath + "\\" + destFile, document);
                    addPicture(programPath + "\\" + destFile, listOfFiles[i].getAbsolutePath(), document);
                } else if (listOfFiles[i].isDirectory()) {
                    getAllFiles(listOfFiles[i].getAbsolutePath(), document);
                }
            }
            //   addString(path + "\\" + pic2, path + "\\" + file2, document); print dir name

        }
    }

    private static void addPicture(String path2File2, String path2Pic2, CustomXWPFDocument document) {
        FileOutputStream fos;
        try {

            fos = new FileOutputStream(new File(path2File2));
            String id;
            try {
                BufferedImage bimg = ImageIO.read(new File(path2Pic2));
                int width = bimg.getWidth();
                int height = bimg.getHeight();
                id = document.addPictureData(new FileInputStream(new File(path2Pic2)), Document.PICTURE_TYPE_JPEG);
                document.createPicture(id, document.getNextPicNameNumber(Document.PICTURE_TYPE_JPEG), width / 2, height / 2);
                document.write(fos);
                fos.flush();
                fos.close();
            } catch (InvalidFormatException ex) {
                Logger.getLogger(Pic2Doc.class.getName()).log(Level.SEVERE, null, ex);
            }

        } catch (FileNotFoundException ex) {
            Logger.getLogger(Pic2Doc.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(Pic2Doc.class.getName()).log(Level.SEVERE, null, ex);
        }

    }

    private static void createTemplateFile() {
        checkExistAndCreate(new File(programPath + "\\Template\\" + template));

    }

    private static void checkExistAndCreate(File f) {
        System.out.println(f.getParentFile());
        if (!f.getParentFile().exists()){
            System.out.println(f.getParentFile());
            f.getParentFile().mkdirs();
        }
        if (!f.exists()) {
            try {
                f.createNewFile();
               
            } catch (IOException ex) {
                Logger.getLogger(Pic2Doc.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }

    private static void createDestFile() {
        File f = new File(programPath + "\\" + destFile);
        if (!f.exists()){
          
            return;
        }
        
        int counter=1;
        String tempDestFile="";
        while (f.exists()){
            
            tempDestFile=destFile.split(".docx")[0]+"("+(counter++)+").docx";
            f=new File(programPath + "\\" + tempDestFile);
        }
          System.out.println(destFile+" is exist, changing to: "+tempDestFile);
        destFile=tempDestFile;
        
    }

    public Pic2Doc(InputStream in) throws IOException {
        super(in);
    }

    public void createPicture(String blipId, int id, int width, int height) {
        final int EMU = 9525;
        width *= EMU;
        height *= EMU;
        //String blipId = getAllPictures().get(id).getPackageRelationship().getId();

        CTInline inline = createParagraph().createRun().getCTR().addNewDrawing().addNewInline();

        String picXml = ""
                + "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"
                + "   <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"
                + "      <pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"
                + "         <pic:nvPicPr>"
                + "            <pic:cNvPr id=\"" + id + "\" name=\"Generated\"/>"
                + "            <pic:cNvPicPr/>"
                + "         </pic:nvPicPr>"
                + "         <pic:blipFill>"
                + "            <a:blip r:embed=\"" + blipId + "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>"
                + "            <a:stretch>"
                + "               <a:fillRect/>"
                + "            </a:stretch>"
                + "         </pic:blipFill>"
                + "         <pic:spPr>"
                + "            <a:xfrm>"
                + "               <a:off x=\"0\" y=\"0\"/>"
                + "               <a:ext cx=\"" + width + "\" cy=\"" + height + "\"/>"
                + "            </a:xfrm>"
                + "            <a:prstGeom prst=\"rect\">"
                + "               <a:avLst/>"
                + "            </a:prstGeom>"
                + "         </pic:spPr>"
                + "      </pic:pic>"
                + "   </a:graphicData>"
                + "</a:graphic>";

        //CTGraphicalObjectData graphicData = inline.addNewGraphic().addNewGraphicData();
        XmlToken xmlToken = null;
        try {
            xmlToken = XmlToken.Factory.parse(picXml);
        } catch (XmlException xe) {
            xe.printStackTrace();
        }
        inline.set(xmlToken);
        //graphicData.set(xmlToken);

        inline.setDistT(0);
        inline.setDistB(0);
        inline.setDistL(0);
        inline.setDistR(0);

        CTPositiveSize2D extent = inline.addNewExtent();
        extent.setCx(width);
        extent.setCy(height);

        CTNonVisualDrawingProps docPr = inline.addNewDocPr();
        docPr.setId(id);
        docPr.setName("Picture " + id);
        docPr.setDescr("Generated");
    }
}
