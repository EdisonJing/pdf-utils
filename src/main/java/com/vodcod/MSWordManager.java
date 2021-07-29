package com.vodcod;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class MSWordManager {
	// word文档
	private Dispatch doc;
	// word运行程序对象
	private ActiveXComponent word;
	// 所有word文档集合
	private Dispatch documents;
	// 选定的范围或插入点
	private Dispatch selection;
	public static int instanceSize = 3;

	// 一个线程存放的MSWordManager数量
	public MSWordManager(int index) {
		if (word == null) {
			word = new ActiveXComponent("Word.Application");
			// 为true表示word应用程序可见
			word.setProperty("Visible", new Variant(false));
		}
		if (documents == null) {
			documents = word.getProperty("Documents").toDispatch();
		}
		if (index == 0) {
			ComThread.InitSTA();
			// 初始化一个线程并放入内存中等待调用
		}
	}

	/**
	 * 打开一个已经存在的文档
	 * 
	 * @param docPath 要打开的文档
	 * @param key     文本框的内容，根据该key获取文本框当前位置
	 * @date 2011-12-9
	 * @author wst
	 */
	public void openDocumentAndGetSelection(String docPath, String key) {
		try {
			closeDocument();// 打开文档
			doc = Dispatch.call(documents, "Open", docPath).toDispatch();
			// shapes集合
			Dispatch shapes = Dispatch.get(doc, "Shapes").toDispatch(); // shape的个数
			String Count = Dispatch.get(shapes, "Count").toString();
			for (int i = 1; i <= Integer.parseInt(Count); i++) {
				// 取得一个shape
				Dispatch shape = Dispatch.call(shapes, "Item", new Variant(i)).toDispatch();
				// 从一个shape里面获取到文本框
				Dispatch textframe = Dispatch.get(shape, "TextFrame").toDispatch();
				boolean hasText = Dispatch.call(textframe, "HasText").toBoolean();
				if (hasText) {
					// 获取该文本框对象
					Dispatch TextRange = Dispatch.get(textframe, "TextRange").toDispatch();
					// 获取文本框中的字符串
					String str = Dispatch.get(TextRange, "Text").toString();
					// 获取指定字符key所在的文本框的位置
					if (str != null && !str.equals("") && str.indexOf(key) > -1) {
						// 当前文本框的位置
						selection = Dispatch.get(textframe, "TextRange").toDispatch();
						// 情况文本框内容
						Dispatch.put(selection, "Text", "");
						break;
					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
			return;
		}
	}

	/**
	 * 在当前位置插入图片
	 * 
	 * @param imagePath 产生图片的路径
	 * @return 成功：true；失败：false
	 */
	public boolean insertImage(String imagePath) {
		try {
			Dispatch.call(Dispatch.get(selection, "InLineShapes").toDispatch(), "AddPicture", imagePath);
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	// 关闭文档
	public void closeDocument() {
		if (doc != null) {
			Dispatch.call(doc, "Close");
			doc = null;
		}
	}

	// 关闭全部应用
	public void close(int index) {
		if (word != null) {
			Dispatch.call(word, "Quit");
			word = null;
		}
		selection = null;
		documents = null;
		if (index == instanceSize) {
			// 释放占用的内存空间，因为com的线程回收不由java的垃圾回收器处理
			ComThread.Release();
		}
	}
}
