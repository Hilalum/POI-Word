import javax.swing.*;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigInteger;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class PoiWord {

	public static void main(String[] args) {
		JFrame jf = new JFrame("DOC�滻");
		jf.setSize(200, 200);
		jf.setLocationRelativeTo(null);
		jf.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);

		JPanel panel = new JPanel();

		// ����һ����ť
		final JButton btn = new JButton("һ���滻");
		final String pathString = "C:\\Users\\dcy\\Desktop\\����\\ģ��\\ģ��.doc";
		final String outPathString = "C:\\Users\\dcy\\Desktop\\����\\���\\���.doc";
		JLabel pathJLabel = new JLabel(pathString, JLabel.LEFT);
		JLabel outPathJLabel = new JLabel(outPathString, JLabel.RIGHT);
		// ��Ӱ�ť�ĵ���¼�������
		btn.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				Map map = new HashMap();
				map.put("-jiafang", "�ְ�");
				map.put("-yifang", "��");
				//replace(map, pathString, outPathString);
				creatTable(map,pathString+"x",outPathString+"x");
			}
		});

		panel.add(btn);
		panel.add(pathJLabel);
		panel.add(outPathJLabel);
		jf.setContentPane(panel);
		jf.setVisible(true);
	}

	public static void replace(Map<String, String> map, String filePath,
			String fileOutPath) {
		HWPFDocument doc = null;

		try {
			InputStream inp = new FileInputStream(filePath);
			POIFSFileSystem fs = new POIFSFileSystem(inp);

			doc = new HWPFDocument(fs);
			Range range = doc.getRange();
			for (Map.Entry<String, String> entry : map.entrySet()) {
				range.replaceText(entry.getKey(), entry.getValue());
			}
			inp.close();
			OutputStream os = new FileOutputStream(fileOutPath);
			doc.write(os);
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	public static void creatTable(Map<String, String> map, String filePath,
			String fileOutPath) {
		try {
			InputStream inp = new FileInputStream(filePath);
			XWPFDocument doc = new XWPFDocument(inp);
			XWPFTable table = doc.createTable(4, 2);
			table.setCellMargins(50, 0, 50, 3000);// top, left, bottom, right
			// table.setInsideHBorder(XWPFBorderType.NONE, 0, 0, "");//ȥ����Ԫ���ĺ���
			table.getRow(0).getCell(0).setText("�ֶ�һ:");
			table.getRow(0).getCell(1).setText("�ֶζ�:");
			table.getRow(1).getCell(0).setText("�ֶ���:");
			table.getRow(1).getCell(1).setText("�ֶ���:");
			inp.close();
			OutputStream os = new FileOutputStream(fileOutPath);
			doc.write(os);
			os.close();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

}
