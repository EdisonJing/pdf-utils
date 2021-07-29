package com.vodcod;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;

public class Word2Pdf2 {
	private static List<File> wordFiles = new ArrayList<>();
	private static ExecutorService pool = Executors.newFixedThreadPool(20);

	public static void listFile(File file) {

		File[] files = file.listFiles();
		if(files!=null) {
			for (File temp : files) {
				// 判断是不是该文件
				if (temp.isDirectory()) {
					// 递归遍历文件夹中是否还有文件夹
					listFile(temp);
				}else {
					if(temp.getName().endsWith(".docx")) {
						System.out.println(temp.getName());
						wordFiles.add(temp);
					}
				}
			}
		}
	}

	public static void main(String args[]) {
		File dir = new File("E:/");
		listFile(dir);

		System.out.println("开始转换...");
		// 开始时间
		long start = System.currentTimeMillis();
		
		CountDownLatch count = new CountDownLatch(wordFiles.size());
		for(File file : wordFiles) {
			String wordFile = file.getAbsolutePath();
			String pdfFile = "D:/cache/pdftest/"+file.getName().substring(0,file.getName().lastIndexOf("."))+".pdf";
			
			pool.execute(new Runnable() {
				@Override
				public void run() {
					ComThread.InitSTA();
					ActiveXComponent app = new ActiveXComponent("Word.Application");
					try {
						// 打开word
						// 设置word不可见,很多博客下面这里都写了这一句话，其实是没有必要的，因为默认就是不可见的，如果设置可见就是会打开一个word文档，对于转化为pdf明显是没有必要的
						// app.setProperty("Visible", false);
						// 获得word中所有打开的文档
						Dispatch documents = app.getProperty("Documents").toDispatch();
						System.out.println("打开文件: " +  wordFile);
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
					} catch (Exception e) {
						e.printStackTrace();
						System.out.println("转换失败" + e.getMessage());
					} finally {
						// 关闭office
						app.invoke("Quit", 0);
						count.countDown();
						ComThread.Release();
					}
				}
			});
		}
		try {
			count.await();
		} catch (InterruptedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		// 结束时间
		long end = System.currentTimeMillis();
		System.out.println("转换成功，用时：" + (end - start) + "ms");

	}
}
