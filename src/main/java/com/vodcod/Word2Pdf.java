package com.vodcod;

import java.io.File;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;

public class Word2Pdf {
	public static void main(String args[]) {
		ActiveXComponent app = null;
		String wordFile = "D:\\cache\\test\\test2.docx";
		String pdfFile = "D:\\cache\\test\\test2.pdf";

		System.out.println("��ʼת��...");
		// ��ʼʱ��
		long start = System.currentTimeMillis();
		try {
			ComThread.InitSTA();
			// ��word
			app = new ActiveXComponent("Word.Application");
			// ����word���ɼ�,�ܶ಩���������ﶼд����һ�仰����ʵ��û�б�Ҫ�ģ���ΪĬ�Ͼ��ǲ��ɼ��ģ�������ÿɼ����ǻ��һ��word�ĵ�������ת��Ϊpdf������û�б�Ҫ��
			//app.setProperty("Visible", false);
			// ���word�����д򿪵��ĵ�
			Dispatch documents = app.getProperty("Documents").toDispatch();
			System.out.println("���ļ�: " + wordFile);
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
			// ����ʱ��
			long end = System.currentTimeMillis();
			System.out.println("ת���ɹ�����ʱ��" + (end - start) + "ms");
		}catch(Exception e) {
			e.printStackTrace();
			System.out.println("ת��ʧ��"+e.getMessage());
		}finally {
			// �ر�office
			app.invoke("Quit", 0);
			ComThread.Release();
		}
	}
}
