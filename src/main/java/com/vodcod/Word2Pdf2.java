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
				// �ж��ǲ��Ǹ��ļ�
				if (temp.isDirectory()) {
					// �ݹ�����ļ������Ƿ����ļ���
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

		System.out.println("��ʼת��...");
		// ��ʼʱ��
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
						// ��word
						// ����word���ɼ�,�ܶ಩���������ﶼд����һ�仰����ʵ��û�б�Ҫ�ģ���ΪĬ�Ͼ��ǲ��ɼ��ģ�������ÿɼ����ǻ��һ��word�ĵ�������ת��Ϊpdf������û�б�Ҫ��
						// app.setProperty("Visible", false);
						// ���word�����д򿪵��ĵ�
						Dispatch documents = app.getProperty("Documents").toDispatch();
						System.out.println("���ļ�: " +  wordFile);
						// ���ĵ�
						Dispatch document = Dispatch.call(documents, "Open", wordFile, false, true).toDispatch();
						// ����ļ����ڵĻ������Ḳ�ǣ���ֱ�ӱ�������������Ҫ�ж��ļ��Ƿ����
						File target = new File(pdfFile);
						if (target.exists()) {
							target.delete();
						}
						System.out.println("���Ϊ: " + pdfFile);
						// ���Ϊ�����ĵ�����Ϊpdf������word����Ϊpdf�ĸ�ʽ���ֵ��17
						Dispatch.call(document, "SaveAs", pdfFile, 17);
						// �ر��ĵ�
						Dispatch.call(document, "Close", false);
					} catch (Exception e) {
						e.printStackTrace();
						System.out.println("ת��ʧ��" + e.getMessage());
					} finally {
						// �ر�office
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
		// ����ʱ��
		long end = System.currentTimeMillis();
		System.out.println("ת���ɹ�����ʱ��" + (end - start) + "ms");

	}
}
