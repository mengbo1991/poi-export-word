package com.censoft.util;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigInteger;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.servlet.ServletContext;
import javax.servlet.http.HttpServletRequest;

import oracle.net.aso.h;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

@Controller
public class WordUtil {

	// log4j日志
	private final Logger logger = LoggerFactory.getLogger(getClass());

	/**
	 * 替换段落里面的变量
	 * 
	 * @param doc
	 *            要替换的文档
	 * @param params
	 *            参数
	 * @param fontFamily  字体
	 * 
	 * @param fontSize 字体大小
	 * 
	 */
	public void replaceInPara(XWPFDocument doc, Map<String, Object> params,String fontFamily,int fontSize) {
		Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
		XWPFParagraph para;
		while (iterator.hasNext()) {
			para = iterator.next();
			this.replaceInPara(para, params,fontFamily,fontSize);
		}
	}

	/**
	 * 替换段落里面的变量
	 * 
	 * @param para
	 *            要替换的段落
	 * @param params
	 *            参数
	 * @param fontFamily  字体
	 * 
	 * @param fontSize 字体大小
	 */
	private void replaceInPara(XWPFParagraph para, Map<String, Object> params,String fontFamily,int fontSize) {
		List<XWPFRun> runs;
		Matcher matcher;
		if (this.matcher(para.getParagraphText()).find()) {
			runs = para.getRuns();
			for (int i = 0; i < runs.size(); i++) {
				XWPFRun run = runs.get(i);
				String runText = run.toString();
				//System.out.println(runText);
				matcher = this.matcher(runText);
				if (matcher.find()) {
					while ((matcher = this.matcher(runText)).find()) {
						runText = matcher.replaceFirst(String.valueOf(params
								.get(matcher.group(1))));
					}
					// 直接调用XWPFRun的setText()方法设置文本时，在底层会重新创建一个XWPFRun，把文本附加在当前文本后面，
					// 所以我们不能直接设值，需要先删除当前run,然后再自己手动插入一个新的run。
					para.removeRun(i);
					XWPFRun newrun = para.insertNewRun(i);
					newrun.setFontSize(fontSize);
					//newrun.setColor("FF0000");
					newrun.setFontFamily(fontFamily);
					newrun.setText(runText);
				}
			}
		}
	}

	/**
	 * 替换表格里面的变量
	 * 
	 * @param doc
	 *            要替换的文档
	 * @param params
	 *            参数
	 * @param fontFamily  字体
	 * 
	 * @param fontSize 字体大小
	 */
	public void replaceInTable(XWPFDocument doc, Map<String, Object> params,String fontFamily,int fontSize) {
		Iterator<XWPFTable> iterator = doc.getTablesIterator();
		XWPFTable table;
		List<XWPFTableRow> rows;
		List<XWPFTableCell> cells;
		List<XWPFParagraph> paras;
		while (iterator.hasNext()) {
			table = iterator.next();
			rows = table.getRows();
			for (XWPFTableRow row : rows) {
				cells = row.getTableCells();
				for (XWPFTableCell cell : cells) {
					paras = cell.getParagraphs();
					for (XWPFParagraph para : paras) {
						this.replaceInPara(para, params,fontFamily,fontSize);
					}
				}
			}
		}
	}
	/***
	 *  修改当前word文档表格里变量的参数属性
	 * @param doc 要修改的文档
	 * @param tableNum 第几个表
	 * @param rowsNum 第几行
	 * @param cellNum 第几列
	 * @param paragraphsNum 第几个段落
	 * @param runsNum 第几个篇幅
	 * @param fontFamily 字体
	 * @param fontSize 大小
	 * @param bold 是否加粗
	 */
	public void updateInTable(XWPFDocument doc, int tableNum, int rowsNum, int cellNum, int paragraphsNum, int runsNum, String fontFamily, String fontSize, boolean bold) {
		List<XWPFTable> allTable = doc.getTables();
		List<XWPFTableRow> rows = allTable.get(tableNum).getRows();
        List<XWPFTableCell> cells = rows.get(rowsNum).getTableCells();
		List<XWPFParagraph> paragraphs = cells.get(cellNum).getParagraphs();
		List<XWPFRun> runs = paragraphs.get(paragraphsNum).getRuns();
		XWPFRun run = runs.get(runsNum);
		fontFamily = StringUtils.trimToEmpty(fontFamily);
		fontSize = StringUtils.trimToEmpty(fontSize);
		if(bold){
			run.setBold(bold);
		}
		if(!"".equals(fontSize)){
			int size = Integer.valueOf(fontSize);
			run.setFontSize(size);
		}
		if(!"".equals(fontFamily)){
			run.setFontFamily(fontFamily);
		}
	}
	/**
	 * 正则匹配字符串
	 * 
	 * @param str
	 * @return
	 */
	private Matcher matcher(String str) {
		Pattern pattern = Pattern.compile("&(.+?)&", Pattern.CASE_INSENSITIVE);
		Matcher matcher = pattern.matcher(str);
		return matcher;
	}
	
	  public static void setCellTextr(XWPFTableCell cell, String text, int width,  
	            boolean isShd, int shdValue, String shdColor,boolean isCenter) {  
	        CTTc cttc = cell.getCTTc();  
	        CTTcPr ctPr = cttc.isSetTcPr() ? cttc.getTcPr() : cttc.addNewTcPr();  
	        CTShd ctshd = ctPr.isSetShd() ? ctPr.getShd() : ctPr.addNewShd();  
	        ctPr.addNewTcW().setW(BigInteger.valueOf(width));  
	        if (isShd) {  
	            /*if (shdValue > 0 && shdValue <= 38) {  
	                ctshd.setVal(STShd.Enum.forInt(shdValue));  
	            }*/  
	        	ctshd.setVal(STShd.CLEAR); 
	            if (shdColor != null) {  
	            	ctshd.setColor("auto");  
	                ctshd.setFill(shdColor);  
	            }  
	        }
	        ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);  
	        if(isCenter){
	        	cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);  
	        }
	        
	       /* CTP ctp = CTP.Factory.newInstance();
			XWPFParagraph p = new XWPFParagraph(ctp, cell);
			p.setAlignment(ParagraphAlignment.CENTER);
	        
			XWPFRun run = p.createRun();
			run.setFontSize(16);
			run.setColor("FF0000");
	        run.setFontFamily("仿宋_GB2312");
	        //run.setText(text);
	        cell.addParagraph(p);
	        //cell.setText(text);
	        
	        run.setText(text);*/
	        //cell.setText(text);  
	        getParagraph(cell,text);
	        
	    }
		/** 
	     * @Description: 跨列合并 
	     */  
	    public void mergeCellsHorizontal(XWPFTable table, int row, int fromCell, int toCell) {  
	        for (int cellIndex = fromCell; cellIndex <= toCell; cellIndex++) {  
	            XWPFTableCell cell = table.getRow(row).getCell(cellIndex);  
	            if ( cellIndex == fromCell ) {  
	                // The first merged cell is set with RESTART merge value  
	                //cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);  
	                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);  
	            } else {  
	                // Cells which join (merge) the first one, are set with CONTINUE  
	                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);  
	            }  
	        }  
	    }  
		 /** 
	     * @Description: 跨行合并 
	     * @see h  ttp://stackoverflow.com/questions/24907541/row-span-with-xwpftable 
	     */  
	    public  void mergeCellsVertically(XWPFTable table, int col, int fromRow, int toRow) {  
	        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {  
	            XWPFTableCell cell = table.getRow(rowIndex).getCell(col);  
	            if ( rowIndex == fromRow ) {  
	                // The first merged cell is set with RESTART merge value  
	                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);  
	            } else {  
	                // Cells which join (merge) the first one, are set with CONTINUE  
	                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);  
	            }  
	        }  
	    }  
		//设置宽度（单元格）
		public void setTableWidth(XWPFTable table,String width){
			   CTTbl ttbl = table.getCTTbl();
			   CTTblPr tblPr = ttbl.getTblPr() == null ? ttbl.addNewTblPr() : ttbl.getTblPr();
			   CTTblWidth tblWidth = tblPr.isSetTblW() ? tblPr.getTblW() : tblPr.addNewTblW();
			   CTJc cTJc=tblPr.addNewJc();
			   cTJc.setVal(STJc.Enum.forString("center"));
			   tblWidth.setW(new BigInteger(width));
			   tblWidth.setType(STTblWidth.DXA);
			}
		//设置单元格（暂时不用）
		public void createTable(XWPFDocument doc,String str) {
			   XWPFTable table = null;
			   table = doc.createTable(1, 1);
			   setTableWidth(table, "9000");
			   XWPFTableRow row = null;
			   row = table.getRow(0);
			   row.setHeight(380);
			   XWPFTableCell cell = null;
			   cell = row.getCell(0);
			   }
		//居中对齐
		public void setCellText(XWPFTableCell cell, String text, int width) {
			   CTTc cttc = cell.getCTTc();
			   CTTcPr cellPr = cttc.addNewTcPr();
			   cellPr.addNewTcW().setW(BigInteger.valueOf(width));
			   CTTcPr ctPr = cttc.addNewTcPr();
			   ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);
			   cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);
			   cell.setText(text);
			}
		//左对齐
		public void setCellTextleft(XWPFTableCell cell, String text, int width) {
			   CTTc cttc = cell.getCTTc();
			   CTTcPr cellPr = cttc.addNewTcPr();
			   cellPr.addNewTcW().setW(BigInteger.valueOf(width));
			   CTTcPr ctPr = cttc.addNewTcPr();
			   ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);
			   cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.LEFT);
			   cell.setText(text);
			}
		public static void getParagraph(XWPFTableCell cell,String cellText){
			CTP ctp = CTP.Factory.newInstance();
			XWPFParagraph p = new XWPFParagraph(ctp, cell);
			p.setAlignment(ParagraphAlignment.CENTER);
	        XWPFRun run = p.createRun();
	        run.setFontSize(14);
	        run.setText(cellText);
	        //run.setBold(true);
	        //run.setFontFamily("仿宋_GB2312");
	        CTRPr rpr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
	        CTFonts fonts = rpr.isSetRFonts() ? rpr.getRFonts() : rpr.addNewRFonts();
	        fonts.setAscii("仿宋_GB2312");
	        fonts.setEastAsia("仿宋_GB2312");
	        fonts.setHAnsi("仿宋_GB2312");
	        cell.setParagraph(p);
		}
		
		/**
		 * @Description 复制一个表格的行带格式,新增到指定位置
		 * @param table 表格
		 * @param sourceRow 需要复制的行
		 * @param rowIndex  复制到第几行
		 * 
		 */
		
		public void copyRow(XWPFTable table,XWPFTableRow sourceRow,int rowIndex){
		    //在表格指定位置新增一行
			XWPFTableRow targetRow = table.insertNewTableRow(rowIndex);
			//复制行属性
			targetRow.getCtRow().setTrPr(sourceRow.getCtRow().getTrPr());
			List<XWPFTableCell> cellList = sourceRow.getTableCells();
			if (null == cellList) {
			    return;
			}
			//复制列及其属性和内容
			XWPFTableCell targetCell = null;
			for (XWPFTableCell sourceCell : cellList) {
			    targetCell = targetRow.addNewTableCell();
			    //列属性
			    targetCell.getCTTc().setTcPr(sourceCell.getCTTc().getTcPr());
			    //段落属性
			    if(sourceCell.getParagraphs()!=null&&sourceCell.getParagraphs().size()>0){                     
			    	targetCell.getParagraphs().get(0).getCTP().setPPr(sourceCell.getParagraphs().get(0).getCTP().getPPr());
		            if(sourceCell.getParagraphs().get(0).getRuns()!=null&&sourceCell.getParagraphs().get(0).getRuns().size()>0){
		            	XWPFRun cellR = targetCell.getParagraphs().get(0).createRun();
		    	        cellR.setText(sourceCell.getText());
		    	        cellR.setBold(sourceCell.getParagraphs().get(0).getRuns().get(0).isBold());
		            }else{
		            	targetCell.setText(sourceCell.getText());
		            }
		        }else{
		        	targetCell.setText(sourceCell.getText());
		        }
		    }
		}

		
}
