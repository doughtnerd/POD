package com.doughtnerd.pod.excel.abstracts;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.TreeMap;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.doughtnerd.pod.excel.exceptions.SheetNotFoundException;

/**
 * This is an abstract class designed to read organized and tabulated data from
 * an xls or an xlsx document and return an <code>ArrayList</code> that contains
 * the data read from a single sheet or a <code>TreeMap</code> that contains the
 * data extracted from the entire document.
 * 
 * @author Christopher Carlson
 *
 * @param <T>
 *            The type of object that is being created through the extraction
 *            process.
 */
public abstract class ExcelReader<T> {

	/**
	 * The workbook the file represents.
	 */
	protected Workbook workbook;

	/**
	 * The original file containing the workbook data.
	 */
	protected File file;

	/**
	 * Creates a new ExcelReader Object.
	 * 
	 * @param file
	 *            The file to process
	 * @throws IOException
	 *             Thrown if an IOException occurred.
	 */
	public ExcelReader(File file) throws IOException {
		FileInputStream fis;
		String extension = FilenameUtils.getExtension(file.getAbsolutePath());
		this.file = file;
		fis = new FileInputStream(file);
		workbook = extension.equals("xls") ? new HSSFWorkbook(fis)
				: extension.equals("xlsx") ? new XSSFWorkbook(fis) : null;
		if (workbook == null) {
			throw new IllegalArgumentException("File needs to be of type: xls or xlsx");
		}
		fis.close();
	}

	/**
	 * Extracts the first row from the specified sheet (assumes headers are in
	 * first row of sheet).
	 * 
	 * @param sheetName
	 *            Name of the sheet to extract headers from.
	 * @return ArrayList of the strings contained in the first row of the sheet.
	 * @throws SheetNotFoundException
	 *             Thrown if the specified sheet was not found.
	 */
	public ArrayList<String> extractHeaders(String sheetName) throws SheetNotFoundException {
		Sheet sheet = this.getSheet(this.workbook, sheetName);
		return extractHeaders(sheet);
	}

	/**
	 * Extracts the first row from the specified sheet (assumes headers are in
	 * first row of sheet).
	 * 
	 * @param sheetIndex
	 *            Index of the sheet to extract headers from.
	 * @return ArrayList of the strings contained in the first row of the sheet.
	 * @throws SheetNotFoundException
	 *             Thrown if the specified sheet was not found.
	 */
	public ArrayList<String> extractHeaders(int sheetIndex) throws SheetNotFoundException {
		Sheet sheet = this.getSheet(this.workbook, sheetIndex);
		return extractHeaders(sheet);
	}

	/**
	 * Extracts the first row from the specified sheet (assumes headers are in
	 * first row of sheet).
	 * 
	 * @param sheet
	 *            Sheet to extract headers from.
	 * @return ArrayList of the strings contained in the first row of the sheet.
	 * @throws SheetNotFoundException
	 *             Thrown if the specified sheet was not found.
	 */
	public ArrayList<String> extractHeaders(Sheet sheet) {
		ArrayList<String> list = new ArrayList<>();
		Row row = sheet.getRow(0);
		Iterator<Cell> iter = row.cellIterator();
		while (iter.hasNext()) {
			list.add(iter.next().toString());
		}
		return list;
	}

	/**
	 * Processes the entire excel document.
	 * 
	 * @param headers
	 *            Whether or not headers are present on every sheet of the file.
	 * @return A TreeMap keyed by sheet name, containing all data T on that
	 *         sheet for the whole workbook.
	 */
	public TreeMap<String, ArrayList<T>> processDocument(boolean headers) {
		TreeMap<String, ArrayList<T>> map = new TreeMap<>();
		int sheetCount = workbook.getNumberOfSheets();
		for (int i = 0; i < sheetCount; i++) {
			String sheetName = workbook.getSheetName(i);
			ArrayList<T> sheetData;
			try {
				sheetData = processSheet(sheetName, headers);
				map.put(sheetName, sheetData);
			} catch (SheetNotFoundException e) {
				e.printStackTrace();
			}
		}
		return map;
	}

	/**
	 * Strips data T from the excel sheet.
	 * 
	 * @param sheetName
	 *            The name of the sheet where the data is found
	 * @param headers
	 *            True if there is a header row present in the sheet, false
	 *            otherwise.
	 * @return An ArrayList containing all T data from the sheet.
	 * @throws SheetNotFoundException
	 *             Thrown if the specified sheetName returned null.
	 */
	public ArrayList<T> processSheet(String sheetName, boolean headers) throws SheetNotFoundException {
		Sheet sheet = getSheet(workbook, sheetName);
		return processSheet(sheet, headers);
	}

	/**
	 * Strips data T from the excel sheet.
	 * 
	 * @param sheetIndex
	 *            The index of the sheet where the data is found.
	 * @param headers
	 *            True if there is a header row present in the sheet, false
	 *            otherwise.
	 * @return An ArrayList containing all T data from the sheet.
	 * @throws SheetNotFoundException
	 *             Thrown if the specified sheetIndex returned null.
	 */
	public ArrayList<T> processSheet(int sheetIndex, boolean headers) {
		Sheet sheet = getSheet(workbook, sheetIndex);
		return processSheet(sheet, headers);
	}

	/**
	 * Strips data T from the excel sheet.
	 * 
	 * @param sheet
	 *            The excel sheet containing the data to extract.
	 * @param headers
	 *            True if there is a header row present in the sheet, false
	 *            otherwise.
	 * @return An ArrayList containing all T data from the sheet.
	 */
	public ArrayList<T> processSheet(Sheet sheet, boolean headers) {
		ArrayList<T> list = new ArrayList<>();
		Iterator<Row> iter = sheet.iterator();
		while (iter.hasNext()) {
			Row row = iter.next();
			if (!headers) {
				T t = extractItem(row);
				if (t != null) {
					list.add(t);
				}
			} else {
				headers = false;
			}
		}
		return list;
	}

	/**
	 * This method tells the reader how to extract data type T from a given row
	 * in the excel sheet.
	 * 
	 * @param row
	 *            The row data is being extracted from
	 * @return
	 */
	protected abstract T extractItem(Row row);

	/**
	 * Scans the workbook for a specific sheet
	 * 
	 * @param workbook
	 *            The workbook to scan
	 * @param sheetName
	 *            The string the sheet should contain
	 * @return The first sheet containing the sheetName given to the method.
	 * @throws SheetNotFoundException
	 *             Thrown if the specified sheetIndex returned null.
	 */
	protected Sheet getSheet(Workbook workbook, String sheetName) throws SheetNotFoundException {
		for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
			Sheet sheet = workbook.getSheetAt(i);
			if (sheet.getSheetName().contains(sheetName)) {
				return sheet;
			}
		}
		throw new SheetNotFoundException("Could not find sheet: " + sheetName);
	}

	/**
	 * Scans the workbook for a specific sheet
	 * 
	 * @param workbook
	 *            The workbook to scan
	 * @param sheetIndex
	 *            The index number of the sheet to grab.
	 * @return The sheet at the given index of the workbook.
	 * @throws SheetNotFoundException
	 *             Thrown if the specified sheetIndex returned null.
	 */
	protected Sheet getSheet(Workbook workbook, int sheetIndex) {
		Sheet sheet = workbook.getSheetAt(sheetIndex);
		return sheet;
	}

	/**
	 * 
	 * @return Iterates through the workbook until it finds a sheet then returns
	 *         that sheet's index within the workbook.
	 */
	public int getFirstSheetIndex() {
		Iterator<Sheet> iter = workbook.sheetIterator();
		while (iter.hasNext()) {
			Sheet sheet = iter.next();
			if (sheet != null) {
				return workbook.getSheetIndex(sheet);
			}
		}
		return -1;
	}

	/**
	 * @return The first visible tab within the workbook.
	 */
	public int getFirstVisibleSheetIndex() {
		return workbook.getFirstVisibleTab();
	}

	/**
	 * Returns a list of sheet names that the workbook contains.
	 * 
	 * @return ArrayList of Strings representing the sheet names this workbook
	 *         contains.
	 */
	public ArrayList<String> getSheetNames() {
		ArrayList<String> list = new ArrayList<>();
		Iterator<Sheet> iter = workbook.sheetIterator();
		while (iter.hasNext()) {
			Sheet sheet = iter.next();
			String name = sheet.getSheetName();
			if (name != null && !name.equals("")) {
				list.add(name);
			}
		}
		return list;
	}

	/**
	 * Iterates through the workbook to find the first sheet whose name contains
	 * the input string.
	 * 
	 * @param string
	 *            The input string the desired sheet should have in its name.
	 * @param caseSensitive
	 *            Whether or not this method should perform a case sensitive
	 *            search.
	 * @return The index of the sheet within the workbook or -1 if no sheet was
	 *         found.
	 */
	public int getFirstIndexOfSheetContaining(String string, boolean caseSensitive) {
		Iterator<Sheet> iter = workbook.sheetIterator();
		while (iter.hasNext()) {
			Sheet sheet = iter.next();
			String name = caseSensitive ? sheet.getSheetName() : sheet.getSheetName().toLowerCase();
			string = caseSensitive ? string : string.toLowerCase();
			if (name.contains(string)) {
				return workbook.getSheetIndex(sheet);
			}
		}
		return -1;
	}

	/**
	 * Converts a map of String to ArrayList of E data to a summarized array of
	 * the E data.
	 * 
	 * @param <T>
	 *            The type of objects that the resulting ArrayList will contain.
	 * @param data
	 *            - The map to extract data from.
	 * @return - A summarized ArrayList of data.
	 */
	public static <T> ArrayList<T> asArrayList(TreeMap<String, ArrayList<T>> data) {
		ArrayList<T> list = new ArrayList<>();
		for (String s : data.keySet()) {
			list.addAll(data.get(s));
		}
		return list;
	}

	/**
	 * <p>
	 * Opens up a file chooser dialog that allows for the selection of Excel
	 * type documents from the user's computer. This method attempts to open the
	 * FileChooser in the specified root directory however, if the dialog is
	 * unable to open the dialog in the specified location, opens to the "" file
	 * directory (usually the user's documents folder).
	 * </p>
	 * <p>
	 * The file types that this chooser will allow are .xls and .xlsx file
	 * types.
	 * </p>
	 * <p>
	 * If interaction == 1 and the user fails to add .xlsx or .xls to the end of
	 * the filename, this method will auto-append .xlsx to the end of the
	 * filename.
	 * </p>
	 * 
	 * @param title
	 *            The title to display in the dialog.
	 * @param root
	 *            The root folder to open the dialog in. If the chooser is
	 *            unable to open the dialog in the specified location, opens to
	 *            the "" file directory (usually the user's documents folder).
	 * @param fileOption
	 *            The available selection options.
	 * @param interaction
	 *            0 to open a file, 1 to save.
	 * @return The file selected by the user or null if no file was selected.
	 */
	public static File getExcelFile(String title, File root, int fileOption, int interaction) {
		JFileChooser chooser = null;
		try {
			chooser = new JFileChooser(root);
		} catch (Exception e) {
			chooser = new JFileChooser(new File(""));
		}
		chooser.setDialogTitle(title);
		chooser.addChoosableFileFilter(
				new FileNameExtensionFilter("MS Excel Workbooks (*.xls, *.xlsx)", "xls", "xlsx"));
		chooser.addChoosableFileFilter(new FileNameExtensionFilter("Excel Workbook (*.xlsx)", "xlsx"));
		chooser.addChoosableFileFilter(new FileNameExtensionFilter("Excel 97-2003 Workbook (*.xls)", "xls"));
		chooser.setAcceptAllFileFilterUsed(false);
		chooser.setFileSelectionMode(fileOption);
		int selection = interaction == 0 ? chooser.showOpenDialog(null) : chooser.showSaveDialog(null);
		if (selection == JFileChooser.APPROVE_OPTION) {
			if (interaction == 1) {
				String path = chooser.getSelectedFile().getAbsolutePath();
				path = path.contains(".xlsx") || path.contains(".xls") ? path : path + ".xlsx";
				return new File(path);
			} else {
				return chooser.getSelectedFile();
			}
		} else {
			return null;
		}
	}

	/**
	 * Tries to get a string value from an excel cell. If the cell is a number,
	 * returns the integer value formatted as a string. If the cell is a string,
	 * returns that cell's string value. If cell==null returns an empty string.
	 * 
	 * @param cell
	 *            The cell to get a string from.
	 * @return A string representing the cell's value.
	 */
	public static String getCellStringValue(Cell cell) {

		if (cell == null) {
			return "";
		}

		String value = "";
		try {
			value = (int) cell.getNumericCellValue() + "";
		} catch (Exception e) {
			value = cell.getStringCellValue();
		}
		return value;
	}
}
