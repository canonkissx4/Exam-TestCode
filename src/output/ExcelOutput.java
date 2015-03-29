package output;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.Serializable;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import commons.CreateExcel;

/**
 * ファイル出力のクラス
 *
 * 花型Excel読込→編集→Output(EXCEL)
 */
public class ExcelOutput implements Serializable {

	/** 選択されたデータ */
	private static List selDatas;
	/** Primefacesとの連携に使うクラス */
	//private StreamedContent file;

	public static void excelFunction() {
		/* チェックボックスで選択されたものを出力するため、未選択、上限越えを判定 */
//		if (selDatas.isEmpty()) {
//			//データ無し。
//		} else if (selDatas.size() > 20) {
//
//		} else {
			// resource file path
			//ExternalContext ec = FacesContext.getCurrentInstance().getExternalContext();

			try {
				// フォーマット読み込み（出力したい帳票ごとにファイル名変更）
				//Input Stream

				InputStream read = new FileInputStream(new File("C:/temp/test.xlsx"));
				OPCPackage pkg = OPCPackage.open(read);

				//low memory? TODO//
//				OPCPackage pkg = OPCPackage.open(new File("C:/temp/test.xlsx"));
				XSSFWorkbook wb = new XSSFWorkbook(pkg);

				// EXCEL作成
				wb = createExcel(wb);

				// 入力データのOUTPUT（重複しないように現在時刻を使用）
				String outputTime = new SimpleDateFormat("yyyyMMddHHmmssSSS").format(new Date());
				FileOutputStream fileOut = new FileOutputStream("C:/temp/" + outputTime + "output.xlsx");
				wb.write(fileOut);
				fileOut.close();

				pkg.close();

//				for(Result data : selDatas){
//						//Data set
//				}

				//log.addKey("帳票");\\
			} catch (InvalidFormatException e) {
				e.printStackTrace();
				//log.setError(e);
			} catch (IOException e) {
				e.printStackTrace();
				//log.setError(e);
			} finally{
				//session.close();
			}
		}
//	}

	private static XSSFWorkbook createExcel(XSSFWorkbook wb) {
		wb.getFontAt((short) 0).setFontName("ＭＳ 明朝");
		wb.getFontAt((short) 0).setFontHeightInPoints((short) 12);
		Sheet sheet = wb.getSheetAt(0);

		// InputSheetBase(sheet);
		Row row = sheet.getRow(0);

		InputSheetDatas(wb, selDatas, 14);

		return wb;
	}

	/**
	 * 各データの帳票入力
	 *
	 * @param wb
	 * @param selDatas
	 * @param start
	 */
	private static void InputSheetDatas(XSSFWorkbook wb, List selDatas, int start) {
		Sheet sheet = wb.getSheetAt(0);
		int startRow = start;
		int formRow = 11;
		Date date;
		String str;

		//align 均等割り付け（インデント）
		short INDENT_CENTER = 7;

		// データ表示の間隔
		float rowHeight = (float) 23;

		startRow = 5;
		Row row;

		row = sheet.createRow(startRow);
		row.setHeightInPoints(rowHeight);

		String val = "TEST テキスト";
		CreateExcel.createCell(wb, row, (short) 0, CellStyle.ALIGN_CENTER, val);




//		for (int i = 0; i < selDatas.size(); i++) {
//			startRow = start + (formRow * i);
//			Row row;
//
//			row = sheet.createRow(startRow);
//			row.setHeightInPoints(rowHeight);
//			// 番号
//			String val = Integer.toString(i + 1);
//			CreateExcel.createCell(wb, row, (short) 0, CellStyle.ALIGN_CENTER, val);
//
//			// テキスト
//			row = sheet.createRow(startRow + 2);
//			row.setHeightInPoints(rowHeight);
//			CreateExcel.createCell(wb, row, (short) 1, INDENT_CENTER, "テキスト");
//			// date = selDatas.get(i).getShibouDate();
////			setDefaultDateForm(wb, row, date);
//			CreateExcel.createCell(wb, row, (short)3, CellStyle.ALIGN_LEFT, selDatas.get(i).getS());
//
//			// 4 テキスト
//			row = sheet.createRow(startRow + 3);
//			row.setHeightInPoints(rowHeight);
//			CreateExcel.createCell(wb, row, (short) 1, INDENT_CENTER, "テキスト");
//			CreateExcel.createCell(wb, row, (short) 3, CellStyle.ALIGN_LEFT, selDatas.get(i).getP());
//			CreateExcel.createCell(wb, row, (short) 6, CellStyle.ALIGN_LEFT, selDatas.get(i).getP() + "  テキスト");
//
//			// 5 テキスト
//			row = sheet.createRow(startRow + 5);
//			row.setHeightInPoints(rowHeight);
//			CreateExcel.createCell(wb, row, (short) 1, INDENT_CENTER, "テキスト");
//			CreateExcel.createCell(wb, row, (short) 3, CellStyle.ALIGN_LEFT, selDatas.get(i).getN());
//
//			// テキスト
//			// テキスト
//			CreateExcel.createCell(wb, row, (short) 18, CellStyle.ALIGN_LEFT, selDatas.get(i).getS());
//
//			// 6 テキスト
//			row = sheet.createRow(startRow + 6);
//			row.setHeightInPoints(rowHeight);
//			CreateExcel.createCell(wb, row, (short) 1, INDENT_CENTER, "テキスト");
//
//			// 7 テキスト
//			row = sheet.createRow(startRow + 7);
//			row.setHeightInPoints(rowHeight);
//			CreateExcel.createCell(wb, row, (short) 1, INDENT_CENTER, "テキスト");
//			CreateExcel.createCell(wb, row, (short) 3, CellStyle.ALIGN_LEFT, selDatas.get(i).getM());
//		}

	}

	public static void main(String[] args) {
		excelFunction();
	}

//	private void nonSelect() {
//		addMessage("確認メッセージ", "項目を選択して下さい。");
//	}
}