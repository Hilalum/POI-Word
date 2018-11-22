import javax.swing.*;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigInteger;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class PoiWord {

	public static void main(String[] args) {
		JFrame jf = new JFrame("DOC替换");
		jf.setSize(400, 500);
		jf.setLocationRelativeTo(null);
		jf.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);

		final JPanel panel = new JPanel();

		// 创建一个按钮
		final JButton btn = new JButton("一键替换");
		final JButton btn1 = new JButton("生成表");
		final String pathString = "C:\\Users\\dcy\\Desktop\\测试\\模版\\模版.docx";
		final String outPathString = "C:\\Users\\dcy\\Desktop\\测试\\输出\\完成.docx";
		JLabel pathJLabel = new JLabel(pathString);
		JLabel outPathJLabel = new JLabel(outPathString);
		// 添加按钮的点击事件监听器
		final Map<String, String> map = new HashMap();
				
				map.put("$(name)","我");
				map.put("$(qq)","爸爸");
		btn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				replace(map, pathString, outPathString);
			}});
		btn1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
			creatTable(map,pathString,outPathString);
		}});
		panel.add(btn);
		panel.add(btn1);
		panel.add(pathJLabel);
		panel.add(outPathJLabel);
		jf.setContentPane(panel);
		jf.setVisible(true);
	}

	public static boolean replace(Map<String, String> map, String filePath,//文档替换
			String fileOutPath) {
		 String[] sp = filePath.split("\\.");
	        String[] dp = fileOutPath.split("\\.");
        if (sp.length <= 0 || dp.length <= 0) {
            return false;
        }
 
        if (
                !sp[sp.length - 1].equalsIgnoreCase("docx")
                        &&
                        !(
                                sp[sp.length - 1].equalsIgnoreCase("doc")
                                        && dp[dp.length - 1].equalsIgnoreCase("doc")
                        )
                ) {
            return false;
        }
 
 
        // 比较文件扩展名
        if (sp[sp.length - 1].equalsIgnoreCase("docx")) {
            XWPFDocument document=null;
			try {
				document = new XWPFDocument(POIXMLDocument.openPackage(filePath));
			} catch (IOException e2) {
				// TODO Auto-generated catch block
				e2.printStackTrace();
			}
            // 替换段落中的指定文字
            Iterator<XWPFParagraph> itPara = document.getParagraphsIterator();
            while (itPara.hasNext()) {
                XWPFParagraph paragraph = itPara.next();
                List<XWPFRun> runs = paragraph.getRuns();
                for (XWPFRun run : runs) {
                	int flag=0;
                    String oneparaString = run.getText(run.getTextPosition());
                    for (int i = 0; i < oneparaString.length(); i++) {
                        if (Character.isWhitespace(oneparaString.charAt(i)) == false) {
                           flag++;
                        }
                    }
                    if ((oneparaString.length()) == 0||oneparaString==null||flag==0){
                        continue;
                    }
                    
                    for (Map.Entry<String, String> entry :
                            map.entrySet()) {
                        oneparaString = oneparaString.replace(entry.getKey(), entry.getValue());
                    }
                    run.setText(oneparaString, 0);
                }
 
            }
 
            // 替换表格中的指定文字
            Iterator<XWPFTable> itTable = document.getTablesIterator();
            while (itTable.hasNext()) {
                XWPFTable table = itTable.next();
                int rcount = table.getNumberOfRows();
                for (int i = 0; i < rcount; i++) {
                    XWPFTableRow row = table.getRow(i);
                    List<XWPFTableCell> cells = row.getTableCells();
                    for (XWPFTableCell cell : cells) {
                        String cellTextString = cell.getText();
                        for (Map.Entry<String, String> e : map.entrySet()) {
                            cellTextString = cellTextString.replace(e.getKey(), e.getValue());
                        }
                        cell.removeParagraph(0);
                        cell.setText(cellTextString);
                    }
                }
            }
            FileOutputStream outStream;
			try {
				outStream = new FileOutputStream(fileOutPath);
            document.write(outStream);
				outStream.close();
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
            return true;
        }
        // doc只能生成doc，如果生成docx会出错
        if ((sp[sp.length - 1].equalsIgnoreCase("doc"))
                && (dp[dp.length - 1].equalsIgnoreCase("doc"))) {
            HWPFDocument document;
			try {
				document = new HWPFDocument(new FileInputStream(filePath));
            Range range = document.getRange();  
            for (Map.Entry<String, String> entry : map.entrySet()) {
                range.replaceText(entry.getKey(), entry.getValue());
            }
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
          
        }
		return false;
	}

   
	public static void creatTable(Map<String, String> map, String filePath,
			String fileOutPath) {
		try {
			InputStream inp = new FileInputStream(filePath);
			XWPFDocument doc = new XWPFDocument(inp);
			XWPFTable table = doc.createTable(4, 2);
			table.setCellMargins(50, 0, 50, 3000);// top, left, bottom, right
			// table.setInsideHBorder(XWPFBorderType.NONE, 0, 0, "");//去除单元格间的横线
			table.getRow(0).getCell(0).setText("字段一:");
			table.getRow(0).getCell(1).setText("字段二:");
			table.getRow(1).getCell(0).setText("字段三:");
			table.getRow(1).getCell(1).setText("字段四:");
			inp.close();
			OutputStream os = new FileOutputStream(fileOutPath);
			doc.write(os);
			os.close();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

}
