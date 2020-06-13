import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;

import java.io.File;

/**
 * @author wangxin
 * @Title
 * @Description
 * @date 2020-06-12 18:18
 */
public class WordToPdfUtil {
    public static void wordToPdf(String wordPath, String pdfPath) {
        ActiveXComponent app = null;
        String wordFile = wordPath;
        File file = new File(wordFile);
        if (file.exists()) {
            String fileName = wordFile.substring(0, wordFile.lastIndexOf("."));
            String pdfFile = pdfPath;
            System.out.println("开始转换...");
            System.out.println("路径为：" + wordPath + "的word开始转换！");
            // 开始时间
            long start = System.currentTimeMillis();
            try {
                //初始化com的线程
                System.out.println("初始化com的线程");
                ComThread.InitSTA();
                System.out.println("word运行程序对象");
                // 打开word
                app = new ActiveXComponent("Word.Application");
                // 设置word不可见,很多博客下面这里都写了这一句话，其实是没有必要的，因为默认就是不可见的，如果设置可见就是会打开一个word文档，对于转化为pdf明显是没有必要的
                //app.setProperty("Visible", false);
                // 获得word中所有打开的文档
                Dispatch documents = app.getProperty("Documents").toDispatch();
                System.out.println("打开文件: " + wordFile);
                // 打开文档
                Dispatch document = Dispatch.call(documents, "Open", wordFile, false, true).toDispatch();
                // 如果文件存在的话，不会覆盖，会直接报错，所以我们需要判断文件是否存在
                File target = new File(pdfFile);
                if (target.exists()) {
                    target.delete();
                }
                System.out.println("另存为: " + pdfFile);
                // 另存为，将文档报错为pdf，其中word保存为pdf的格式宏的值是17
                Dispatch.call(document, "SaveAs", pdfFile, 17);
                // 关闭文档
                Dispatch.call(document, "Close", false);

                // 结束时间
                long end = System.currentTimeMillis();
                System.out.println("转换成功，用时：" + (end - start) + "ms");

                System.out.println("路径为：" + wordPath + "的word转换成功，用时：" + (end - start) + "ms");

                //删除掉原始word
     		  /* File file2 = new File(wordPath);
     		   if(file2.exists()){
     			   file2.delete();
     		   }*/
            } catch (Exception e) {
                System.out.println("转换失败" + e.getMessage());
                System.out.println("路径为：" + wordPath + "的word转换失败，捕获异常：" + e.getMessage());
            } finally {
                // 关闭office
                app.invoke("Quit", 0);
                System.out.println("转换路径为：" + wordPath + "的word进程关闭");

                //关闭com的线程
                ComThread.Release();
            }
        } else {
            System.out.println("文件不存在！");
            System.out.println("路径为：" + wordPath + "的word文件不存在!");
        }
    }

    public static void main(String[] args) {
        String word = "D:\\test\\word\\was搭建部署手册.docx";
        String name = "搭建部署手册".concat(".pdf");
        String pdf = "D:\\test\\word\\" + name;
        wordToPdf(word, pdf);
    }
}
