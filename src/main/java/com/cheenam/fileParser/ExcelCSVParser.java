package com.cheenam.fileParser;

import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.opencsv.CSVReader;

public class ExcelCSVParser<D> {
	public static class ParseBuilder<D> {
		private RowExtractor<D> converter;
		private char delimiter = CSV_DELIM;
		private boolean hasHeader;
		private int sheets;

		public ParseBuilder() {
		}

		public ExcelCSVParser<D> build() {
			return new ExcelCSVParser<D>(this);
		}

		public ParseBuilder<D> converter(RowExtractor<D> converter) {
			this.converter = converter;
			return this;
		}

		public ParseBuilder<D> csvDelimiter(char delimiter) {
			this.delimiter = delimiter;
			return this;
		}

		public ParseBuilder<D> sheets(int sheetCount) {
			this.sheets = sheetCount;
			return this;
		}

		public ParseBuilder<D> withHeader() {
			this.hasHeader = true;
			return this;
		}

	}

	public interface RowExtractor<D> {
		D convert(Object[] row);
	}

	private static final char CSV_DELIM = ',';

	private static final int XL_SHEET_COUNT = 1;

	public static <D> ParseBuilder<D> builder(Class<D> cls) {
		return new ParseBuilder<D>();
	}

	private ParseBuilder<D> info;

	private ExcelCSVParser(ParseBuilder<D> info) {
		this.info = info;
	}

	private D extractObject(Iterator<Row> rowIterator) {
		Row row = rowIterator.next();
		Iterator<Cell> cellIterator = row.cellIterator();
		Object[] rowVals = new Object[row.getLastCellNum()];
		while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
			rowVals[cell.getColumnIndex()] = getValue(cell);
		}
		return info.converter.convert(rowVals);
	}

	private void extractSheet(Sheet sheet, List<D> objList) {
		Iterator<Row> rowIterator = sheet.iterator();
		if (rowIterator.hasNext() && info.hasHeader) {
			rowIterator.next();
		}
		while (rowIterator.hasNext()) {
			D obj = extractObject(rowIterator);
			objList.add(obj);
		}
	}

	private Object getValue(Cell cell) {
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_STRING:
			return cell.getStringCellValue();
		case Cell.CELL_TYPE_NUMERIC:
			return cell.getNumericCellValue();
		case Cell.CELL_TYPE_BOOLEAN:
			return cell.getBooleanCellValue();
		case Cell.CELL_TYPE_ERROR:
			return cell.getErrorCellValue();
		case Cell.CELL_TYPE_FORMULA:
			return cell.getCellFormula();
		case Cell.CELL_TYPE_BLANK:
			return null;
		}
		return null;
	}

	private boolean isExcel(InputStream is) throws Exception {

		return POIXMLDocument.hasOOXMLHeader(is) /* .xlsx */
				|| POIFSFileSystem.hasPOIFSHeader(is); /* .xls */
	}

	private List<D> parseCSV(InputStream in) throws Exception {
		List<D> objList = new ArrayList<>();
		InputStreamReader isr = new InputStreamReader(in);
		try (CSVReader cvsr = new CSVReader(isr, info.delimiter)) {
			List<String[]> allRows = cvsr.readAll();
			int start = info.hasHeader ? 1 : 0;
			for (int i = start; i < allRows.size(); i++) {
				D obj = info.converter.convert(allRows.get(i));
				objList.add(obj);
			}
		}
		return objList;
	}

	private List<D> parseExcel(InputStream is) throws Exception {
		Workbook workbook = WorkbookFactory.create(is);
		int sheetCount = Math.min(workbook.getNumberOfSheets(), XL_SHEET_COUNT);
		List<D> objList = new ArrayList<>();
		sheetCount = (info.sheets == 0) ? sheetCount : info.sheets;
		for (int i = 0; i < sheetCount; i++) {
			Sheet sheet = workbook.getSheetAt(i);
			extractSheet(sheet, objList);
		}
		return objList;
	}

	public List<D> read(InputStream is) throws Exception {
		List<D> objList = null;
		try (BufferedInputStream buf = new BufferedInputStream(is)) {
			if (isExcel(buf)) { // XLSX, XLS
				objList = parseExcel(buf);
			} else { // CSV
				objList = parseCSV(buf);
			}
		}
		return objList;
	}

	public List<D> read(String fileName) throws Exception {
		try (FileInputStream is = new FileInputStream(fileName)) {
			return read(is);
		}
	}
}
