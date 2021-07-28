package com.vodcod;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.xwpf.usermodel.TextAlignment;

//needed jars: fr.opensagres.poi.xwpf.converter.core-2.0.1.jar, 
//             fr.opensagres.poi.xwpf.converter.pdf-2.0.1.jar,
//             fr.opensagres.xdocreport.itext.extension-2.0.1.jar,
//             itext-2.1.7.jar                                  

//needed jars: apache poi and it's dependencies
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.poi.xwpf.usermodel.XWPFTableCell.XWPFVertAlign;

import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;

public class DOCXToPDFConverterSampleMin {

	public static void main(String[] args) throws Exception {

		String docPath = "D:\\cache\\test\\test.docx";
		String pdfPath = "D:\\cache\\test\\test.pdf";

		InputStream in = new FileInputStream(new File(docPath));
		XWPFDocument document = new XWPFDocument(in);
		setFontType(document);
		PdfOptions options = PdfOptions.create();
		OutputStream out = new FileOutputStream(new File(pdfPath));
		PdfConverter.getInstance().convert(document, out, options);

		document.close();
		out.close();

	}

	private static void setFontType(XWPFDocument document) {
		//转换文档中文字字体
		List<XWPFParagraph> paragraphs = document.getParagraphs();
		if(paragraphs != null && paragraphs.size()>0){
			for (XWPFParagraph paragraph : paragraphs) {
				List<XWPFRun> runs = paragraph.getRuns();
				System.out.println(paragraph.isPageBreak());
				if(runs !=null && runs.size()>0){
					for (XWPFRun run : runs) {
						run.setFontFamily("宋体");
					}
				}
			}
		}
		//转换表格里的字体 我也不想俄罗斯套娃但是不套真不能设置字体
		List<XWPFTable> tables = document.getTables();
		for (XWPFTable table : tables) {
			List<XWPFTableRow> rows = table.getRows();
			for (XWPFTableRow row : rows) {
				List<XWPFTableCell> tableCells = row.getTableCells();
				for (XWPFTableCell tableCell : tableCells) {
					tableCell.setVerticalAlignment(XWPFVertAlign.CENTER);
					List<XWPFParagraph> paragraphs1 = tableCell.getParagraphs();
					for (XWPFParagraph xwpfParagraph : paragraphs1) {
						xwpfParagraph.setVerticalAlignment(TextAlignment.CENTER);
						List<XWPFRun> runs = xwpfParagraph.getRuns();
						for (XWPFRun run : runs) {
							run.setFontFamily("宋体");
						}
					}
				}
			}
		}

	}
}
