package com.com;

import com.UserMapper;
import com.baomidou.mybatisplus.core.assist.ISqlRunner;
import com.domain.User;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.hslf.extractor.PowerPointExtractor;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.openxmlformats.schemas.drawingml.x2006.main.CTRegularTextRun;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextBody;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraph;
import org.openxmlformats.schemas.presentationml.x2006.main.CTGroupShape;
import org.openxmlformats.schemas.presentationml.x2006.main.CTShape;
import org.openxmlformats.schemas.presentationml.x2006.main.CTSlide;
import org.springframework.beans.factory.annotation.Autowired;

import java.io.*;
import java.util.List;

/**
 * @author
 * @date 2021年10月07日19:58
 */
public class DocRead {
    /**
     * @Description: POI 读取  word
     * @create: 2019-07-27 9:48
     * @update logs
     * @throws Exception
     */

    @Autowired
    public  UserMapper userMapper;


    private static int maxx = 10000;
    String filepath = "C:\\Users\\lenovo\\Desktop\\111\\readFile";
    //判断编码格式方法
    private static  String get_code(File sourceFile) {
        String charset = "GBK";
        byte[] first3Bytes = new byte[3];
        try {
            boolean checked = false;
            BufferedInputStream bis = new BufferedInputStream(new FileInputStream(sourceFile));
            bis.mark(0);
            int read = bis.read(first3Bytes, 0, 3);
            if (read == -1) {
                bis.close();
                return charset; //文件编码为 ANSI
            } else if (first3Bytes[0] == (byte) 0xFF
                    && first3Bytes[1] == (byte) 0xFE) {
                charset = "UTF-16LE"; //文件编码为 Unicode
                checked = true;
            } else if (first3Bytes[0] == (byte) 0xFE
                    && first3Bytes[1] == (byte) 0xFF) {
                charset = "UTF-16BE"; //文件编码为 Unicode big endian
                checked = true;
            } else if (first3Bytes[0] == (byte) 0xEF
                    && first3Bytes[1] == (byte) 0xBB
                    && first3Bytes[2] == (byte) 0xBF) {
                charset = "UTF-8"; //文件编码为 UTF-8
                checked = true;
            }
            bis.reset();
            if (!checked) {
                int loc = 0;
                while ((read = bis.read()) != -1) {
                    loc++;
                    if (read >= 0xF0)
                        break;
                    if (0x80 <= read && read <= 0xBF) // 单独出现BF以下的，也算是GBK
                        break;
                    if (0xC0 <= read && read <= 0xDF) {
                        read = bis.read();
                        if (0x80 <= read && read <= 0xBF) // 双字节 (0xC0 - 0xDF)
                            // (0x80
                            // - 0xBF),也可能在GB编码内
                            continue;
                        else
                            break;
                    } else if (0xE0 <= read && read <= 0xEF) {// 也有可能出错，但是几率较小
                        read = bis.read();
                        if (0x80 <= read && read <= 0xBF) {
                            read = bis.read();
                            if (0x80 <= read && read <= 0xBF) {
                                charset = "UTF-8";
                                break;
                            } else
                                break;
                        } else
                            break;
                    }
                }
            }
            bis.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return charset;

    }

    //    @SuppressWarnings("resource")
    public static String readWord(String filePath) throws Exception{
        if(filePath.equals("")) return null;
//        List<String> linList = new ArrayList<String>();
        String buffer = "";
        try {
            if (filePath.endsWith(".doc")) {
                InputStream fis = new FileInputStream(new File(filePath));
                WordExtractor ex = new WordExtractor(fis);
                buffer = ex.getText();
                fis.close();
                ex.close();

            } else if (filePath.endsWith(".docx")) {
                FileInputStream fis = new FileInputStream(filePath);
                XWPFDocument xdoc = new XWPFDocument(fis);
                XWPFWordExtractor ex = new XWPFWordExtractor(xdoc);
                buffer = ex.getText();
                ex.close();
                fis.close();
                xdoc.close();
//                return buffer;
            }
            else if(filePath.endsWith(".pdf")) {
                PDDocument ex;
                InputStream fis = new FileInputStream(new File(filePath));
                ex = PDDocument.load(fis);
                PDFTextStripper stripper = new PDFTextStripper();
                buffer = stripper.getText(ex);
                fis.close();
                ex.close();
            }
            else if(filePath.endsWith(".txt")) {
                File file = new File(filePath);
                String code = get_code(file);
                System.out.println("code: " + code);
//                code = "UTF-8";
                InputStream is = new FileInputStream(file);
                InputStreamReader isr = new InputStreamReader(is, code);
                BufferedReader fis = new BufferedReader(isr);
                String linetxt = null;
                //result用来存储文件内容
                StringBuilder sb = new StringBuilder();
                //按使用readLine方法，一次读一行
                while ((linetxt = fis.readLine()) != null && sb.length() < maxx) {
                    System.out.println(linetxt);
                    sb.append(linetxt);
                    sb.append(" ");
                }
                is.close();
                isr.close();
                fis.close();
                buffer = sb.toString();
//                System.out.println("tex\n" + buffer);
            }
//            --------
/**
            else if(filePath.endsWith("xls") || filePath.endsWith("xlsx")) {
                StringBuilder sb = new StringBuilder();
                FileInputStream fis = new FileInputStream(filePath);
                Workbook wb = null;   //Workbook 不能close 关闭fis即可
                if(filePath.endsWith("xsl")) {
                    wb = new HSSFWorkbook(fis);
                }
                else {
                    wb = new XSSFWorkbook(fis);
                }
                for(int sheetIndex = 0; sheetIndex < wb.getNumberOfSheets() && sb.length() < maxx; sheetIndex++) {
                    Sheet sheet = wb.getSheetAt(sheetIndex);     //读取sheet 0

                    int firstRowIndex = sheet.getFirstRowNum();   //设置变量的第一行
                    int lastRowIndex = sheet.getLastRowNum();     //设置变量的行
                    System.out.println("firstRowIndex: "+firstRowIndex);
                    System.out.println("lastRowIndex: "+lastRowIndex);

                    for(int rIndex = firstRowIndex; rIndex <= lastRowIndex && sb.length() < maxx; rIndex++) {   //遍历行
                        System.out.println("rIndex: " + rIndex);
                        Row row = sheet.getRow(rIndex);
                        if (row != null) {
                            int firstCellIndex = row.getFirstCellNum();
                            int lastCellIndex = row.getLastCellNum();
                            System.out.println("1c: " + firstCellIndex + "lc: " + lastCellIndex);
                            for (int cIndex = firstCellIndex; cIndex < lastCellIndex && sb.length() < maxx; cIndex++) {   //遍历列
                                Cell cell = row.getCell(cIndex);
                                System.out.println(cell);
                                if (cell != null) {
                                    sb.append(cell.toString());
                                    sb.append(" ");
//                                    System.out.println(cell.toString());
                                }
                            }
                        }
                    }
                }
                fis.close();
                buffer = sb.toString();
            }
*/
//            --------
//            --------
            else if(filePath.endsWith("ppt")) {
                FileInputStream fis = new FileInputStream(new File(filePath));
                PowerPointExtractor ex=new PowerPointExtractor(fis);
                buffer = ex.getText();
                fis.close();
                ex.close();
            }
            else if(filePath.endsWith("pptx")) {
                StringBuilder sb = new StringBuilder();
                FileInputStream fis = new FileInputStream(filePath);
                XMLSlideShow xmlSlideShow = new XMLSlideShow(fis);
                List<XSLFSlide> slides = xmlSlideShow.getSlides();
                for(XSLFSlide slide:slides){
                    CTSlide rawSlide = slide.getXmlObject();
                    CTGroupShape gs = rawSlide.getCSld().getSpTree();
                    CTShape[] shapes = gs.getSpArray();
                    for(CTShape shape:shapes){
                        CTTextBody tb = shape.getTxBody();
                        if(null==tb){
                            continue;
                        }
                        CTTextParagraph[] paras = tb.getPArray();
                        for(CTTextParagraph textParagraph:paras){
                            CTRegularTextRun[] textRuns = textParagraph.getRArray();
                            for(CTRegularTextRun textRun:textRuns){
                                sb.append(textRun.getT() + " ");
                            }
                        }
                    }
                }
                buffer = sb.toString();
                xmlSlideShow.close();
                fis.close();
            }
            else {
                return null;
            }
            buffer = buffer.replace("\n|\r", " ");
//            buffer = buffer.replace("'", " ");
            if(buffer.length() > maxx) buffer = buffer.substring(0,maxx);
            return buffer;
        } catch (Exception e) {
            System.out.print("error---->"+filePath);
            e.printStackTrace();
            return null;
        }
    }

    public static void main(String[] args) throws Exception {
//        String txt = DocRead.readWord("C:\\Users\\lenovo\\Desktop\\111\\readFile\\123.txt");
//        String docx = DocRead.readWord("C:\\Users\\lenovo\\Desktop\\111\\readFile\\123.docx");
//        String pdf = DocRead.readWord("C:\\Users\\lenovo\\Desktop\\111\\readFile\\Hive.pdf");
//        String ppt = DocRead.readWord("C:\\Users\\lenovo\\Desktop\\111\\readFile\\spark.ppt");

//        System.out.println("txt = " + txt );
//        System.out.println("docx = " + docx );
//        System.out.println("pdf = " + pdf );
//        System.out.println("ppt = " + ppt );

        System.out.println(("----- selectAll method test ------"));
        List<User> userList = userMapper.selectList(null);
        for(User user:userList) {
            System.out.println(user);
        }
    }
}