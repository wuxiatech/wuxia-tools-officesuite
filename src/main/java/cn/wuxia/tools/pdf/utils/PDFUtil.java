package cn.wuxia.tools.pdf.utils;

import cn.wuxia.common.util.ArrayUtil;
import cn.wuxia.common.util.ListUtil;
import cn.wuxia.common.util.StringUtil;
import cn.wuxia.tools.word.utils.WordUtil;
import com.aspose.words.Document;
import com.aspose.words.SaveFormat;
import com.google.common.collect.Lists;
import com.itextpdf.text.BadElementException;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Element;
import com.itextpdf.text.Image;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;
import com.itextpdf.text.pdf.parser.ImageRenderInfo;
import com.itextpdf.text.pdf.parser.PdfReaderContentParser;
import com.itextpdf.text.pdf.parser.RenderListener;
import com.itextpdf.text.pdf.parser.TextRenderInfo;
import org.apache.commons.lang3.ArrayUtils;
import org.springframework.core.io.ClassPathResource;

import java.io.*;
import java.net.URI;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;

public class PDFUtil {
    static {
        WordUtil.isWordLicense();
    }

    public static void main(String[] args) {
        doc2pdf("/Users/songlin/Downloads/广州市长护定点机构初评表（10000032-稻铭患者).doc", "/Users/songlin/Downloads/广州市长护定点机构初评表（10000032-稻铭患者).pdf");
        FileOutputStream os = null;
        try {
            File file = new File("/Users/songlin/Downloads/广州市长护定点机构初评表（10000032-稻铭患者)2.pdf");
            os = new FileOutputStream(file);
            printSign(
                    new FileInputStream(
                            new File("/Users/songlin/Downloads/广州市长护定点机构初评表（10000032-稻铭患者).pdf")
                    ), os,
                    new FileInputStream(
                            new File("/Users/songlin/app/dmidea/aixin-admin/src/main/webapp/resources/images/gongzhang.png")
                    ), 118f, 118f, Lists.newArrayList("公章"));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    public static void doc2pdf(String inPath, String outPath) {
        FileOutputStream os = null;
        try {
            File file = new File(outPath);
            os = new FileOutputStream(file);
            Document doc = new Document(inPath);
            doc.save(os, SaveFormat.PDF);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (os != null) {
                try {
                    os.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    public static void doc2pdf(InputStream inputStream, OutputStream outputStream) {
        try {
            Document doc = new Document(inputStream);
            doc.save(outputStream, SaveFormat.PDF);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void doc2pdf(File inputFile, OutputStream outputStream) {
        try {
            Document doc = new Document(new FileInputStream(inputFile));
            doc.save(outputStream, SaveFormat.PDF);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


//    /**
//     * 生成 PDF 文件
//     *
//     * @param out  输出流
//     * @param html HTML字符串
//     * @throws IOException       IO异常
//     * @throws DocumentException Document异常
//     */
//    public static void html2pdf(String html, OutputStream out) throws IOException, com.lowagie.text.DocumentException {
//        org.xhtmlrenderer.pdf.ITextRenderer renderer = new org.xhtmlrenderer.pdf.ITextRenderer();
//        renderer.setDocumentFromString(html);
//        // 解决中文支持问题
////        ITextFontResolver fontResolver = renderer.getFontResolver();
////        fontResolver.addFont("pdf/font/fangsong.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
////        fontResolver.addFont("pdf/font/PingFangSC.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
//        renderer.layout();
//        renderer.createPDF(out);
//    }
//
//
//    /**
//     * 生成 PDF 文件
//     *
//     * @param out  输出流
//     * @param file HTML文件
//     * @throws IOException       IO异常
//     * @throws DocumentException Document异常
//     */
//    public static void html2pdf(File file, OutputStream out) throws IOException, com.lowagie.text.DocumentException {
//        org.xhtmlrenderer.pdf.ITextRenderer renderer = new org.xhtmlrenderer.pdf.ITextRenderer();
//        renderer.setDocument(file);
//        // 解决中文支持问题
//        org.xhtmlrenderer.pdf.ITextFontResolver fontResolver = renderer.getFontResolver();
//        fontResolver.addFont("pdf/font/fangsong.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
////        fontResolver.addFont("pdf/font/PingFangSC.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
//        renderer.layout();
//        renderer.createPDF(out);
//    }


    //
//
//    /**
//     *
//     * @Title: insertWatermarkText
//     * @Description: PDF生成水印
//     * @author mzl
//     * @param doc
//     * @param watermarkText
//     * @throws Exception
//     * @throws
//     */
//    private static void insertWatermarkText(Document doc, String watermarkText) throws Exception
//    {
//
//        Shape watermark = new Shape(doc, ShapeType.TEXT_PLAIN_TEXT);
//
//
//        //水印内容
//        watermark.getTextPath().setText(watermarkText);
//        //水印字体
//        watermark.getTextPath().setFontFamily("宋体");
//        //水印宽度
//        watermark.setWidth(500);
//        //水印高度
//        watermark.setHeight(100);
//        //旋转水印
//        watermark.setRotation(-40);
//        //水印颜色
//        watermark.getFill().setColor(Color.lightGray);
//        watermark.setStrokeColor(Color.lightGray);
//
//        watermark.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);
//        watermark.setRelativeVerticalPosition(RelativeVerticalPosition.PAGE);
//        watermark.setWrapType(WrapType.NONE);
//        watermark.setVerticalAlignment(VerticalAlignment.CENTER);
//        watermark.setHorizontalAlignment(HorizontalAlignment.CENTER);
//
//        Paragraph watermarkPara = new Paragraph(doc);
//        watermarkPara.appendChild(watermark);
//
//        for (Section sect : doc.getSections())
//        {
//            insertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HEADER_PRIMARY);
//            insertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HEADER_FIRST);
//            insertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HEADER_EVEN);
//        }
//        System.out.println("Watermark Set");
//    }
//
//
//
//    private static void insertWatermarkIntoHeader(Paragraph watermarkPara, Section sect, int headerType) throws Exception
//    {
//        HeaderFooter header = sect.getHeadersFooters().getByHeaderFooterType(headerType);
//
//        if (header == null)
//        {
//            header = new HeaderFooter(sect.getDocument(), headerType);
//            sect.getHeadersFooters().add(header);
//        }
//
//        header.appendChild(watermarkPara.deepClone(true));
//    }
    public static void printSign(InputStream inputStream, OutputStream outputStream, InputStream pageHeaderInputstream, float sealWidth,
                                 float sealHeight, List<String> keyWords) {
        //支持多关键字，默认选择第一个找到的关键字
        PdfReader pdfReader;
        PdfStamper pdfStamper = null;
        try {
            pdfReader = new PdfReader(inputStream);
            pdfStamper = new PdfStamper(pdfReader, outputStream);
            List<List<float[]>> arrayLists = findKeywords(keyWords, pdfReader);//查找关键字所在坐标
            //一个坐标也没找到，就返回
            if (ListUtil.isEmpty(arrayLists)) {
                return;
            }
            Image pageHeaderImg = getImgByInputstream(pageHeaderInputstream);
            pageHeaderImg.setAlignment(Element.ALIGN_LEFT);
            pageHeaderImg.scaleAbsolute(sealWidth, sealHeight);// 控制签章大小

            if (!ListUtil.isEmpty(arrayLists.get(0))) {
                for (int i = 0; i < arrayLists.get(0).size(); i++) {
                    PdfContentByte overContent = pdfStamper.getOverContent((int) arrayLists.get

                            (0).get(i)[2]);

                    pageHeaderImg.setAbsolutePosition(arrayLists.get(0).get(i)[0],

                            arrayLists.get(0).get(i)[1] - sealHeight / 2);// 控制图片位置
                    overContent.addImage(pageHeaderImg);//将图片加入pdf的内容中
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            //此处一定要关闭流，否则可能会出现乱码
            if (pdfStamper != null) {
                try {
                    pdfStamper.close();
                } catch (com.itextpdf.text.DocumentException e) {
                    e.printStackTrace();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    public static void printSign(OutputStream outputStream) {
        List<String> keyWords = new ArrayList<String>();
        keyWords.add("公章");//支持多关键字，默认选择第一个找到的关键字
        PdfReader pdfReader;
        PdfStamper pdfStamper = null;
        try {
            pdfReader = new PdfReader(((ByteArrayOutputStream) outputStream).toByteArray());
            pdfStamper = new PdfStamper(pdfReader, outputStream);
            List<List<float[]>> arrayLists = findKeywords(keyWords, pdfReader);//查找关键字所在坐标
            //一个坐标也没找到，就返回
            if (ListUtil.isEmpty(arrayLists)) {
                return;
            }
            if (!ListUtil.isEmpty(arrayLists.get(0))) {
                for (int i = 0; i < arrayLists.get(0).size(); i++) {
                    PdfContentByte overContent = pdfStamper.getOverContent((int) arrayLists.get

                            (0).get(i)[2]);
                    String imgPath = "/resource/lodop/sign.png";
                    float sealWidth = 150f;
                    float sealHeight = 95f;
                    InputStream pageHeaderInputstream = PDFUtil.class.getResourceAsStream

                            (imgPath);
                    Image pageHeaderImg = getImgByInputstream(pageHeaderInputstream);
                    pageHeaderImg.setAlignment(Element.ALIGN_LEFT);
                    pageHeaderImg.scaleAbsolute(sealWidth, sealHeight);// 控制签章大小
                    pageHeaderImg.setAbsolutePosition(arrayLists.get(0).get(i)[0],

                            arrayLists.get(0).get(i)[1] - sealHeight / 2);// 控制图片位置
                    overContent.addImage(pageHeaderImg);//将图片加入pdf的内容中
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            //此处一定要关闭流，否则可能会出现乱码
            if (pdfStamper != null) {
                try {
                    pdfStamper.close();
                } catch (com.itextpdf.text.DocumentException e) {
                    e.printStackTrace();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    /**
     * 根据关键字返回对应的坐标
     *
     * @param keyWords
     * @param pdfReader
     * @return
     */
    private static List<List<float[]>> findKeywords(final List<String> keyWords, PdfReader pdfReader) {
        if (keyWords == null || keyWords.size() == 0) {
            return null;
        }
        int pageNum = pdfReader.getNumberOfPages();
        final List<List<float[]>> arrayLists = new ArrayList<List<float[]>>(keyWords.size());
        for (int k = 0; k < keyWords.size(); k++) {
            List<float[]> positions = new ArrayList<float[]>();
            arrayLists.add(positions);
        }
        PdfReaderContentParser pdfReaderContentParser = new PdfReaderContentParser(pdfReader);
        try {
            for (int i = 1; i <= pageNum; i++) {
                final int finalI = i;
                pdfReaderContentParser.processContent(i, new RenderListener() {
                    private StringBuilder pdfsb = new StringBuilder();
                    private float yy = -1f;

                    @Override
                    public void renderText(TextRenderInfo textRenderInfo) {
                        String text = textRenderInfo.getText();
                        com.itextpdf.awt.geom.Rectangle2D.Float boundingRectange =

                                textRenderInfo.getBaseline().getBoundingRectange();
                        if (yy == -1f) {
                            yy = boundingRectange.y;
                        }
                        if (yy != boundingRectange.y) {
                            yy = boundingRectange.y;
                            pdfsb.setLength(0);
                        }
                        pdfsb.append(text);
                        if (pdfsb.length() > 0) {
                            for (int j = 0; j < keyWords.size(); j++) {
                                String[] key_words = StringUtil.split(keyWords.get(j), ",");
                                //假如配置了多个关键字，找到一个就跑
                                for (final String key_word : key_words) {
//                                    if (arrayLists.get(j) != null) {
//                                        break;
//                                    }
                                    if (pdfsb.length() > 0 && pdfsb.toString

                                            ().contains(key_word)) {
                                        float[] resu = new float[3];
                                        resu[0] = boundingRectange.x +
                                                boundingRectange.width - 50;
                                        resu[1] = boundingRectange.y;
                                        resu[2] = finalI;
                                        arrayLists.get(j).add(resu);
                                        pdfsb.setLength(0);
                                        break;
                                    }
                                }
                            }
                        }
                    }

                    @Override
                    public void renderImage(ImageRenderInfo arg0) {
                        //renderImage
                    }

                    @Override
                    public void endTextBlock() {
                        //endTextBlock
                    }

                    @Override
                    public void beginTextBlock() {
                        //beginTextBlock
                    }
                });
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return arrayLists;
    }

    public static List<String> parseList(String source, String regex) {
        if (source == null || "".equals(source)) {
            return null;
        }
        List<String> strList = new ArrayList<String>();
        if (regex == null || "".equals(regex)) {
            strList.add(source);
        } else {
            String[] strArr = source.split(regex);
            for (String str : strArr) {
                if (str != null || !"".equals(str)) {
                    strList.add(str);
                }
            }
        }
        return strList;
    }

    private static Image getImgByInputstream(InputStream is) {
        if (is == null) {
            return null;
        }
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        Image img = null;
        try {
            readInputStream(is, output);
            try {
                img = Image.getInstance(output.toByteArray());
            } catch (BadElementException e) {
                e.printStackTrace();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return img;
    }

    public static void readInputStream(InputStream inputStream, OutputStream outputStream) throws IOException {
        byte[] buffer = new byte[2048];
        int n = 0;
        while (-1 != (n = inputStream.read(buffer))) {
            outputStream.write(buffer, 0, n);
        }
        inputStream.close();
    }


}
