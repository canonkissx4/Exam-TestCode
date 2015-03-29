package commons;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateExcel {

	private XSSFWorkbook wrkBook;
	private XSSFSheet wrkSheet;

	/**
	 * コンストラクタ
	 */
	public CreateExcel() {

	}

//	public static void main(String[] args) {
//		CreateExcel xls = new CreateExcel();
//		// xls.makeExcel("/kppsoudan/WebContent/resources/demo/excel/写真.xls");
//		System.out.println("start");
//		System.out.println("end");
//	}
//
//	public void makeExcel(String filePath) {
//		try {
//			// ファイル読み込み
//			wrkBook = CreateExcel.openBook(filePath);
//			wrkSheet = wrkBook.getSheetAt(0);
//
//			// 描画オブジェクトの生成
//			XSSFDrawing pat = wrkSheet.createDrawingPatriarch();
//
//			// 指定セルと結合最終セルを取得
//			XSSFCell cellS = getCell(wrkSheet, 14, 1);
//			XSSFCell cellE = getRegionLastCell(wrkSheet, cellS);
//
//			// ファイル出力
//			wrkBook.write(new FileOutputStream("./test.xls"));
//		} catch (Exception e) {
//			System.out.println("■err makeExcel");
//			System.out.println(e.toString());
//		}
//
//	}

//	/**
//	 * ブックの取得
//	 *
//	 * @return HSSFWorkbook
//	 */
//	public static XSSFWorkbook openBook(String path) {
//		XSSFWorkbook book = null;
//		try {
//			FileInputStream readFile = new FileInputStream(path);
//			book = new XSSFWorkbook(readFile);
//		} catch (Exception e) {
//			System.out.println(e.toString());
//		}
//		return book;
//	}

//	/**
//	 * 画像ファイルの読み込み
//	 *
//	 * @param filePath
//	 * @return byte配列
//	 */
//	public byte[] readImage(String filePath) {
//
//		byte[] imgBytes = null;
//		try {
//			imgBytes = IOUtils.toByteArray(new FileInputStream(filePath));
//		} catch (Exception e) {
//			System.out.println("画像の読込に失敗");
//		}
//		return imgBytes;
//	}
//
//	/**
//	 * 指定セルが結合されている場合、の右下のセルを取得
//	 *
//	 */
//	public XSSFCell getRegionLastCell(XSSFSheet sht, XSSFCell cell) {
//		XSSFCell nextCell = null;
//		System.out.println("getRegionLastCell start");
//		try {
//			// 結合セルを全て取得
//			System.out.println("  size:" + sht.getNumMergedRegions());
//			for (int i = 0; i < sht.getNumMergedRegions(); i++) {
//				CellRangeAddress range = sht.getMergedRegion(i);
//
//				if (range.isInRange(cell.getRowIndex(), cell.getColumnIndex())) {
//					// 結合セルが指定セルと一致した場合、結合セルの右下のセルを取得
//					System.out.println("■getRegionLastCell col:"
//							+ range.getLastColumn() + " row:"
//							+ range.getLastRow());
//					nextCell = getCell(sht, range.getLastRow(),
//							range.getLastColumn());
//					break;
//				}
//			}
//		} catch (Exception e) {
//			System.out.println("■err getRegionLastCell col:"
//					+ cell.getColumnIndex() + " row:" + cell.getRowIndex());
//		}
//		return nextCell;
//	}

	/**
	 * シートとセル位置を指定することでセルを取得
	 *
	 * @param sht
	 * @param row
	 * @param line
	 * @return
	 */
	public XSSFCell getCell(XSSFSheet sht, int row, int col) {

		XSSFCell cell = null;
		try {
			// セルの取得
			cell = sht.getRow(row).getCell(col);
		} catch (Exception e) {
			System.out.println("■err sht:" + sht.getSheetName() + " row:" + row
					+ " col:" + col);
			System.out.println(e.toString());
		}

		return cell;

	}

	/**
	 * セルの作成（フォーマット依存・文字列）
	 *
	 * @param row
	 * @param column
	 *            カラム番号
	 * @param value
	 *            入力の文字(STRING)
	 */
	public static void createCell(Row row, short column, String value) {
		Cell cell = row.getCell(column);
		if (cell != null) {
			cell.setCellValue(value);
		}
	}

	/**
	 * セルの作成（フォーマット依存・数値）
	 *
	 * @param row
	 * @param column
	 *            カラム番号
	 * @param value
	 *            入力の文字(Integer)
	 */
	public static void createCell(Row row, short column, Integer value) {
		Cell cell = row.getCell(column);
		cell.setCellValue((double) value);
	}

	/**
	 * セルの作成・取得（フォーマット依存・数値）<br>
	 * ※ExcelCreatorではExcelを開かないと関数を再計算しない為、<br>
	 * 　計算した値を取得後、フォーマットの関数を削除する。
	 *
	 * @param row
	 * @param column
	 *            カラム番号
	 * @return 関数にて計算した値(int)
	 */
	public static int createCell(Row row, short column) {
		Cell cell = row.getCell(column);
		// 関数を削除
		cell.setCellFormula(null);
		return (int) cell.getNumericCellValue();
	}

	/**
	 * セルの作成(新規作成必要セルの場合）<br>
	 *
	 * @param wb
	 * @param row
	 * @param column
	 *            カラム番号
	 * @param halign
	 *            配置
	 * @param value
	 *            入力の文字(STRING)
	 */
	public static void createCell(XSSFWorkbook wb, Row row, short column,
			short halign, String value) {
		Cell cell = row.createCell(column);
		cell.setCellValue(value);
		CellStyle cellStyle = wb.createCellStyle();
		cellStyle.setAlignment(halign);
		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cell.setCellStyle(cellStyle);
	}

}
