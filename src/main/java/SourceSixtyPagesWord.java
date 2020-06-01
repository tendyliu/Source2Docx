import cn.hutool.core.io.FileUtil;
import cn.hutool.core.lang.Console;
import cn.hutool.core.util.StrUtil;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.ObjectFactory;

import java.io.File;
import java.io.FileFilter;
import java.util.ArrayList;
import java.util.List;

/**
 * 扫描指定目录的java文件，按照软件著作权要求， 生成两个word文档，一个是从最开始的2000行代码， 一个是从最后开始的2000行代码
 * 不包含空行和注释行信息
 */
public class SourceSixtyPagesWord {

    public static void main(String[] args) throws Exception {

        SourceSixtyPagesWord ssw = new SourceSixtyPagesWord();

        if(args.length<4 || args.length>4){
            Console.print("usage:  java -jar Source2Docx-1.0.jar  sourceDir  fileType  beginDocName  endDocName");
            Console.print(StrUtil.CRLF);
            Console.print("example:  java -jar Source2Docx-1.0.jar  /git/projectA  .java  begin.docx end.docx");
            Console.print(StrUtil.CRLF);
            return;
        }

        ssw.scanAndGenerateSourceDoc(args[0], args[1], args[2], args[3]);
        Console.print("Docxs are generated.");
        Console.print(StrUtil.CRLF);
    }

    /**
     * 随着源码长度以及样式的差异，不能保证60页精确文档，因此采用生成各超过30页的两个文档，再进行人工裁剪页数手动合并的模式。
     * @param sourceDir  要扫描源码的目录
     * @param fileType  扫描的文件类型后缀，可以是 ".java", ".js"等，目前仅支持一种类型
     * @param beginDoc  要保存的开头源码的Word文档
     * @param endDoc    要保存的结束源码的Word文档
     * @throws Exception
     */
    public void scanAndGenerateSourceDoc(String sourceDir, String fileType, String beginDoc, String endDoc) throws Exception {
        List<String> beginSource2000 = new ArrayList<>();
        List<String> endSource2000 = new ArrayList<>();
        scanSource(sourceDir, fileType, beginSource2000, endSource2000);
        genereteWordDoc(beginDoc, beginSource2000);
        genereteWordDoc(endDoc, endSource2000);
    }

    public void genereteWordDoc(String fileName, List<String> sourceLines)throws Exception {

        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();

        // Create the main document part (word/document.xml)
        MainDocumentPart wordDocumentPart = new MainDocumentPart();

        // Create main document part content
        ObjectFactory factory = Context.getWmlObjectFactory();
        org.docx4j.wml.Body body = factory.createBody();
        org.docx4j.wml.Document wmlDocumentEl = factory.createDocument();
        wmlDocumentEl.setBody(body);

        // Put the content in the part
        wordDocumentPart.setJaxbElement(wmlDocumentEl);

        // Add the main document part to the package relationships
        // (creating it if necessary)
        wordMLPackage.addTargetPart(wordDocumentPart);

        for(int i=0; i<sourceLines.size(); i++){
            wordDocumentPart.addParagraphOfText(sourceLines.get(i));
        }

        wordMLPackage.save(new File(fileName) );
    }

    //扫描指定目录及其子目录， 找到指定文件类型的文件清单列表，首先从第一个文件开始读取行信息，忽略掉空行以及注释行，读取2000行代码信息
    //随后从最后一个文件开始读取行信息， 忽略掉空行以及注释行， 读取2000行代码信息
    //如果所有文件的代码行数达不到数量，则报错抛出。
    public void scanSource(String sourceDir, String fileType, List<String> beginSource2000, List<String> endSource2000){
        FileFilter fileFilter = new FileFilter() {
            @Override
            public boolean accept(File pathname) {
                if(pathname.getName().endsWith(fileType)){
                    return true;
                }
                return false;
            }
        };

        List<File> fileList = FileUtil.loopFiles(sourceDir, fileFilter);
        System.out.println("scaned "+fileList.size()+" files");

        fileList.sort((a,b)->a.getAbsolutePath().compareTo(b.getAbsolutePath()));

        //开始读取头2000行代码信息
        for(int i=0; i<fileList.size();i++){
            File file = fileList.get(i);
            List<String> lines = FileUtil.readLines(file, "UTF-8");
            for(String line : lines){
                String trimLine = StrUtil.trimToEmpty(line);
                if(isSourceLine(trimLine)){
                    beginSource2000.add(line);
                    if(beginSource2000.size()>=2000){
                        break;
                    }
                }
            }
            if(beginSource2000.size()>=2000){
                break;
            }
        }


        //开始读取结束的2000行代码信息
        for(int i=fileList.size()-1; i>=0;i--){
            File file = fileList.get(i);
            List<String> lines = FileUtil.readLines(file, "UTF-8");
            for(String line : lines){
                String trimLine = StrUtil.trimToEmpty(line);
                if(isSourceLine(trimLine)){
                    endSource2000.add(line);
                    if(endSource2000.size()>=2000){
                        break;
                    }
                }
            }
            if(endSource2000.size()>=2000){
                break;
            }
        }
    }

    public boolean isSourceLine (String lineText){
        if(StrUtil.isEmpty(lineText)){
            return false;
        }

        //可以带注释
//        if(StrUtil.startWithAny(lineText, "//", "*", "/*")){
//            return false;
//        }
        return true;
    }
}
