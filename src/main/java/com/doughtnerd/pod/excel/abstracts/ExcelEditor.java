package com.doughtnerd.pod.excel.abstracts;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.doughtnerd.pod.excel.exceptions.SheetNotFoundException;

/**
 * This class reads in an excel file and allows for the user to edit specific
 * parts of a given row. After editing, this editors workbook can be saved to a
 * file location on disk.
 * 
 * @author Christopher Carlson
 *
 */
public abstract class ExcelEditor {

	/**
	 * The workbook the file represents.
	 */
	protected Workbook workbook;

	/**
	 * The original file containing the workbook data.
	 */
	protected File file;

	/**
	 * Creates a new ExcelEditor Object.
	 * 
	 * @param file
	 *            The file to process.
	 * @throws IOException
	 *             Thrown if an IOException occurred.
	 */
	public ExcelEditor(File file) throws IOException {
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
	 * Starts the processing of the excel sheet.
	 * 
	 * @param sheetName
	 *            The name of the sheet where the data is found
	 * @param headers
	 *            True if there is a header row present in the sheet, false
	 *            otherwise.
	 * @throws SheetNotFoundException
	 *             Thrown if the specified sheetIndex returned null.
	 */
	public void processSheet(String sheetName, boolean headers) throws SheetNotFoundException {
		Sheet sheet = getSheet(workbook, sheetName);
		processSheet(sheet, headers);
	}

	/**
	 * Starts the processing of the excel sheet.
	 * 
	 * @param sheetIndex
	 *            The index of the sheet where the data is found
	 * @param headers
	 *            True if there is a header row present in the sheet, false
	 *            otherwise.
	 * @throws SheetNotFoundException
	 *             Thrown if the specified sheetIndex returned null.
	 */
	public void processSheet(int sheetIndex, boolean headers) throws SheetNotFoundException {
		Sheet sheet = getSheet(workbook, sheetIndex);
		processSheet(sheet, headers);
	}

	/**
	 * Processes the given sheet.
	 * 
	 * @param sheet
	 *            The sheet to process.
	 * @param headers
	 *            True if there are headers present in the first row, false
	 *            otherwise.
	 */
	public void processSheet(Sheet sheet, boolean headers) {
		if (sheet != null) {
			Iterator<Row> iter = sheet.iterator();
			while (iter.hasNext()) {
				Row current = iter.next();
				if (!headers) {
					editRow(current);
				} else {
					headers = false;
				}
			}
		}
	}

	/**
	 * Instructs the Excel Editor how to edit the given row.
	 * 
	 * @param row
	 *            The current Excel row being edited.
	 */
	protected abstract void editRow(Row row);

	/**
	 * Scans the workbook for a specific sheet
	 * 
	 * @param workbook
	 *            The workbook to scan
	 * @param sheetName
	 *            The string the sheet should contain
	 * @return The first sheet containing the sheetName given to the method
	 * @throws SheetNotFoundException
	 *             Thrown if the specified sheetName returned null.
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
	 * Returns the sheet at the given index.
	 * 
	 * @param workbook
	 *            The workbook to scan
	 * @param sheetIndex
	 *            The index number of the sheet to retrieve.
	 * @return The sheet at the given index of the workbook.
	 * @throws SheetNotFoundException
	 *             Thrown if the specified sheetIndex returned null.
	 */
	protected Sheet getSheet(Workbook workbook, int sheetIndex) throws SheetNotFoundException {
		Sheet sheet = workbook.getSheetAt(sheetIndex);
		if (sheet == null) {
			throw new SheetNotFoundException("Could not find sheet at index: " + sheetIndex);
		}
		return sheet;
	}

	/**
	 * Saves this ExcelEditor objects workbook to the specified file.
	 * 
	 * @param file
	 *            The file location to write to.
	 * @throws IOException
	 *             Thrown if a file access error occurred.
	 */
	public void save(File file) throws IOException {
		FileOutputStream out = new FileOutputStream(file);
		this.workbook.write(out);
		out.close();
	}
}
