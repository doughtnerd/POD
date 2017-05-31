package com.doughtnerd.pod.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Objects;
import java.util.ArrayList;
import java.util.Date;
import java.util.TreeMap;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.doughtnerd.pod.excel.abstracts.ExcelRowObject;
import com.doughtnerd.pod.excel.enums.ExcelFileType;

/**
 * This class provides static methods to handle writing operations of excel
 * sheets and workbooks.
 * 
 * @author Christopher Carlson
 *
 */
public final class ExcelWriter {

	/**
	 * The default sheet name used in various sheet-generating methods if no
	 * valid sheet name is provided.
	 */
	public static final String DEFAULT_SHEETNAME = "Sheet1";

	/**
	 * Writes the given data to the given workbook on a sheet named after the
	 * given sheetName using the given headers.
	 * 
	 * @param <T>
	 *            The type of objects being written to the sheet. All objects
	 *            must extend ExcelRowObject.
	 * @param workbook
	 *            The workbook to write the sheet to.
	 * @param sheetName
	 *            The name the sheet should have in the workbook.
	 * @param headers
	 *            The list of headers for the sheet.
	 * @param data
	 *            The List of ExcelObjects to write to the file.
	 */
	public static <T extends ExcelRowObject> void writeNewSheetToWorkbook(Workbook workbook, String sheetName,
			List<String> headers, List<T> data) {
		Sheet sheet = workbook.createSheet(sheetName);
		int startRow = 0;
		if (headers != null) {
			writeHeaders(headers, sheet, startRow++);
		}
		writeData(workbook, data, sheet, startRow);
	}

	/**
	 * Writes a TreeMap of String sheetName to List&lt;? extends ExcelObject&gt;
	 * pairings to a workbook. Each key in the map represents a sheet and the
	 * value represents the data that the new sheet should contain.
	 * 
	 * @param workbook
	 *            The workbook to write the mapped data to.
	 * @param data
	 *            The map of data to write to the workbook.
	 */
	public static void writeMapToWorkook(Workbook workbook,
			TreeMap<String, ? extends List<? extends ExcelRowObject>> data) {
		for (String s : data.keySet()) {
			writeNewSheetToWorkbook(workbook, s, null, data.get(s));
		}
	}

	/**
	 * Writes a TreeMap of String sheetName to List&lt;? extends ExcelObject&gt;
	 * pairings to a workbook. Each key in the map represents a sheet and the
	 * value represents the data that the new sheet should contain.
	 * 
	 * @param workbook
	 *            The workbook to write the mapped data to.
	 * @param headers
	 *            The headers to use for each sheet.
	 * @param data
	 *            The map of data to write to the workbook.
	 */
	public static void writeMapToWorkook(Workbook workbook, List<String> headers,
			TreeMap<String, ? extends List<? extends ExcelRowObject>> data) {
		for (String s : data.keySet()) {
			writeNewSheetToWorkbook(workbook, s, headers, data.get(s));
		}
	}

	/**
	 * <p>
	 * Writes a TreeMap of String sheetName to List&lt;E extends ExcelObject&gt;
	 * pairings to a workbook. Each key in the map represents a sheet and the
	 * value represents the data that the new sheet should contain.
	 * </p>
	 * <p>
	 * The headers must contain the same amount of key-value pairings as the
	 * data and every key in the data must be contained in the headers
	 * otherwise, an IllegalArgumentException will be thrown by this method.
	 * </p>
	 * 
	 * @param workbook
	 *            The workbook to write the mapped data to.
	 * @param headers
	 *            The headers to use for each sheet.
	 * @param data
	 *            The map of data to write to the workbook.
	 */
	public static void writeMapToWorkook(Workbook workbook, TreeMap<String, ? extends List<String>> headers,
			TreeMap<String, ? extends List<? extends ExcelRowObject>> data) {
		if (headers.size() != data.size()) {
			throw new IllegalArgumentException("Headers must have the same amount of key-value pairings as the data.");
		}
		if (!headers.keySet().equals(data.keySet())) {
			throw new IllegalArgumentException("Headers must contain the same key values as the data.");
		}
		for (String s : data.keySet()) {
			writeNewSheetToWorkbook(workbook, s, headers.get(s), data.get(s));
		}
	}

	/**
	 * Writes a workbook to given path and filename. If the given path does not
	 * match the type of workbook being written, writes to a file with the
	 * proper extension instead.
	 * 
	 * @param workbook
	 *            The workbook to write to file.
	 * @param path
	 *            The destination of the workbook file.
	 * @throws IOException
	 *             Thrown if access to the file was not allowed.
	 */
	public static void writeWorkbookToFile(Workbook workbook, String path) throws IOException {
		File file = correctFileExtension(workbook, path);
		FileOutputStream out = new FileOutputStream(file);
		workbook.write(out);
		out.close();
		workbook.close();
		workbook = null;
		System.out.println("Excel written successfully...");
	}

	/**
	 * Ensures that the file that the workbook will be written to has the
	 * correct file extension. If it doesn't replaces the existing extension
	 * with the correct one.
	 * 
	 * @param workbook
	 *            The workbook that is being written to file.
	 * @param path
	 *            The path of the file being written to.
	 * @return A File representing the write location of the workbook with the
	 *         proper file extension.
	 */
	private static File correctFileExtension(Workbook workbook, String path) {
		String ext = FilenameUtils.getExtension(path);
		if (workbook instanceof HSSFWorkbook) {
			if (!ext.equals("xls")) {
				path = path.replace(ext, "xls");
			}
		} else {
			if (!ext.equals("xlsx")) {
				path = path.replace(ext, "xlsx");
			}
		}
		return new File(path);
	}

	/**
	 * This method writes a List of ExcelObjects to a new Sheet in a new
	 * Workbook then returns the workbook to the user.
	 * 
	 * @deprecated Instead, use
	 *             {@link #writeNewSheetToNewWorkbook(ExcelFileType, String, List, List)}
	 * 
	 * @param <T>
	 *            The type of data being written to the sheet. Must extend
	 *            ExcelRowObject.
	 * @param workbookType
	 *            Either xlsx or xls. If null or empty, defaults to xlsx.
	 * @param sheetName
	 *            The name the sheet should have in the new workbook. If null or
	 *            empty, defaults to DEFAULT_SHEETNAME.
	 * @param headers
	 *            List of String headers that that the sheet should use. If null
	 *            or empty, no headers are written to the sheet.
	 * @param data
	 *            List of Objects which extend ExcelObject that will be written
	 *            as rows to the sheet.
	 * @return The workbook the data was written to.
	 */
	public static <T extends ExcelRowObject> Workbook writeNewSheetToNewWorkbook(String workbookType, String sheetName,
			List<String> headers, List<T> data) {
		if (data == null || data.size() == 0) {
			throw new IllegalArgumentException("There was no data.");
		}
		if (sheetName == null || sheetName.equals("")) {
			sheetName = DEFAULT_SHEETNAME;
		}
		Workbook workbook = getNewWorkbook(workbookType);
		Sheet sheet = workbook.createSheet(sheetName);
		int startRow = 0;
		if (headers != null && headers.size() != 0) {
			writeHeaders(headers, sheet, startRow++);
		}
		writeData(workbook, data, sheet, startRow);
		return workbook;
	}

	/**
	 * This method writes a List of ExcelObjects to a new Sheet in a new
	 * Workbook then returns the workbook to the user.
	 * 
	 * @param <T>
	 *            The type of data being written to the sheet. Must extend
	 *            ExcelRowObject.
	 * @param workbookType
	 *            A valid ExcelFileType enum value. If none is supplied,
	 *            defaults to XLSX.
	 * @param sheetName
	 *            The name the sheet should have in the new workbook. If null or
	 *            empty, defaults to DEFAULT_SHEETNAME.
	 * @param headers
	 *            List of String headers that that the sheet should use. If null
	 *            or empty, no headers are written to the sheet.
	 * @param data
	 *            List of Objects which extend ExcelObject that will be written
	 *            as rows to the sheet.
	 * @return The workbook the data was written to.
	 */
	public static <T extends ExcelRowObject> Workbook writeNewSheetToNewWorkbook(ExcelFileType workbookType,
			String sheetName, List<String> headers, List<T> data) {
		if (data == null || data.size() == 0) {
			throw new IllegalArgumentException("There was no data");
		}
		if (sheetName == null || sheetName.equals("")) {
			sheetName = DEFAULT_SHEETNAME;
		}
		Workbook workbook = getNewWorkbook(workbookType);
		Sheet sheet = workbook.createSheet(sheetName);
		int startRow = 0;
		if (headers != null && headers.size() != 0) {
			writeHeaders(headers, sheet, startRow++);
		}
		writeData(workbook, data, sheet, startRow);
		return workbook;
	}

	/**
	 * Helper method that writes headers to the sheet.
	 * 
	 * @param headers
	 *            The list of headers to write to the sheet.
	 * @param sheet
	 *            The sheet to write the headers to.
	 */
	private static void writeHeaders(List<String> headers, Sheet sheet, int startRow) {
		Row headerRow = sheet.createRow(startRow++);
		for (int i = 0; i < headers.size(); i++) {
			Cell head = headerRow.createCell(i);
			setCellValue(headers.get(i), head);
		}
	}

	/**
	 * Helper method that writes data to an excel sheet with formats.
	 * 
	 * @param workbook
	 *            The workbook to write the sheet to.
	 * @param data
	 *            The data to write to the sheet.
	 * @param allFormatTypes
	 *            The format types that exist across all of the data.
	 * @param sheet
	 *            The sheet that all data is written to and added to the
	 *            workbook.
	 */
	private static <T extends ExcelRowObject> void writeData(Workbook workbook, List<T> data, Sheet sheet,
			int startRow) {
		ArrayList<CellStyle> addedStyles = new ArrayList<>();
		System.out.println("Writing data to: " + sheet.getSheetName());
		for (T key : data) {
			Row row = sheet.createRow(startRow++);
			ExcelCellObject[] objArr = key.toCellObjectArray();
			Objects.requireNonNull(objArr, "ExcelRowObject.toCellObjectArray() cannot result in a null object");
			int cellnum = 0;
			for (ExcelCellObject obj : objArr) {
				Cell cell = row.createCell(cellnum++);
				if (obj != null) {
					Object value = obj.getValue();

					setCellFormat(workbook, cell, addedStyles, obj);
					setCellValue(value, cell);
				}
			}
		}
	}

	/**
	 * If the workbook being processed is an instance of SXSSFWorkbook, this
	 * method will flush the rows in the given sheet to disk in an attempt to
	 * free system memory.
	 * 
	 * @param workbook
	 *            The workbook object being written to, used to check for
	 *            SXSSFWorkbook instance.
	 * @param sheet
	 *            The sheet whose data should be flushed.
	 */
	@SuppressWarnings("unused")
	private static void tryFlushRows(Workbook workbook, Sheet sheet) {
		if (workbook instanceof SXSSFWorkbook) {
			try {
				((SXSSFSheet) sheet).flushRows();
				System.out.println("Flushed Sheet to Disk");
			} catch (Exception e) {
				e.printStackTrace();
			}
		} else {
			System.out.println("Workbook not instance of SXSSF, Cannot flush to disk");
		}
	}

	/**
	 * This method is responsible for setting the style (formats) of each cell
	 * that is placed into the document.
	 * 
	 * @param workbook
	 *            The workbook that will contain the styling.
	 * @param cell
	 *            The cell that is to have styling added.
	 * @param addedStyles
	 *            The list of currently existing styles that have been added to
	 *            the workbook.
	 * @param obj
	 *            The ExcelCellObject being added to the cell.
	 */
	private static void setCellFormat(Workbook workbook, Cell cell, ArrayList<CellStyle> addedStyles,
			ExcelCellObject obj) {
		if (!styleListContains(addedStyles, obj)) {
			setNewFormat(workbook, cell, addedStyles, obj);
		} else {
			setExistingFormat(cell, addedStyles, obj);
		}
	}

	/**
	 * Sets the cell format to an existing cell format, assuming the format can
	 * be found in the cell styles.
	 * 
	 * @param cell
	 *            The cell to add the style to.
	 * @param addedStyles
	 *            The list of current styles that exist in the workbook.
	 * @param obj
	 *            The ExcelCellObject whose formatting properties need to be
	 *            matched.
	 */
	private static void setExistingFormat(Cell cell, ArrayList<CellStyle> addedStyles, ExcelCellObject obj) {
		for (CellStyle s : addedStyles) {
			if (s.getDataFormatString().equals(obj.getDataFormat().getFormatString())
					&& s.getWrapText() == obj.getWrapText()) {
				cell.setCellStyle(s);
				break;
			}
		}
	}

	/**
	 * Sets the cell format to a new format type and adds that format data to
	 * addedFormats and addedStyles.
	 * 
	 * @param workbook
	 *            The workbook that will contain the new cell style.
	 * @param cell
	 *            Cell to add the style to.
	 * @param addedStyles
	 *            The list of styles that currently exist in the workbook.
	 * @param obj
	 *            The ExcelCellObject whose formats are being used to format the
	 *            cell.
	 */
	private static void setNewFormat(Workbook workbook, Cell cell, ArrayList<CellStyle> addedStyles,
			ExcelCellObject obj) {
		CellStyle format = workbook.createCellStyle();
		DataFormat formatObj = workbook.createDataFormat();
		format.setDataFormat(formatObj.getFormat(obj.getDataFormat().getFormatString()));
		format.setWrapText(obj.getWrapText());
		cell.setCellStyle(format);
		addedStyles.add(format);
	}

	/**
	 * Helper method that searches through a list of styles to see if there is a
	 * match between the format string and the style.
	 * 
	 * @param styles
	 *            List of styles to search through.
	 * @param formatString
	 *            The String to match off of.
	 * @return A boolean representing whether or not a match was found in the
	 *         list.
	 */
	private static boolean styleListContains(ArrayList<CellStyle> addedStyles, ExcelCellObject obj) {
		for (CellStyle c : addedStyles) {
			if (c.getDataFormatString().equals(obj.getDataFormat().getFormatString())
					&& c.getWrapText() == obj.getWrapText()) {
				return true;
			}
		}
		return false;
	}

	/**
	 * Creates a new HSSFWorkbook (xls), XSSFWorkbook (xlsx), or SXSSFWorkbook
	 * (streamable xlsx) depending on the type passed to this method.
	 * 
	 * @deprecated Instead, use {@link #getNewWorkbook(ExcelFileType)}
	 * 
	 * @param type
	 *            The type of workbook to create. If type is null, empty, or
	 *            type does not equal xls, xlsx, or streamable-xlsx (shortened
	 *            to sxlsx) defaults to xlsx.
	 * @return The new xls, xlsx, or streamable-xlsx type workbook.
	 */
	public static Workbook getNewWorkbook(String type) {
		if (type == null || type.equals("") || (!type.equals("xls") && !type.equals("xlsx") && !type.equals("sxlsx"))) {
			type = "xlsx";
		}
		Workbook workbook = type.equals("xls") ? new HSSFWorkbook()
				: type.equals("xlsx") ? new XSSFWorkbook() : type.equals("sxlsx") ? new SXSSFWorkbook() : null;
		return workbook;
	}

	/**
	 * Creates a new HSSFWorkbook (xls), XSSFWorkbook (xlsx), or SXSSFWorkbook
	 * (streamable xlsx) depending on the type passed to this method.
	 * 
	 * @param type
	 *            The type of workbook to create. Defaults to xlsx.
	 * @return The new xls, xlsx, or streamable-xlsx type workbook.
	 */
	public static Workbook getNewWorkbook(ExcelFileType type) {
		switch (type) {
		case XLS:
			return new HSSFWorkbook();
		case SXLSX:
			return new SXSSFWorkbook();
		default:
			return new XSSFWorkbook();
		}
	}

	/**
	 * Looks at the object that is going into the cell, sets the cell value type
	 * accordingly, and adds the object to the cell.
	 * 
	 * @param obj
	 *            The object going in the cell.
	 * @param cell
	 *            The cell that will contain the object.
	 */
	private static void setCellValue(Object obj, Cell cell) {
		if (obj instanceof Date)
			cell.setCellValue((Date) obj);
		else if (obj instanceof Boolean)
			cell.setCellValue((Boolean) obj);
		else if (obj instanceof String || obj instanceof Character)
			cell.setCellValue((String) obj);
		else if (obj instanceof Double)
			cell.setCellValue((Double) obj);
		else if (obj instanceof Integer)
			cell.setCellValue((double) ((Integer) obj).intValue());
	}
}
