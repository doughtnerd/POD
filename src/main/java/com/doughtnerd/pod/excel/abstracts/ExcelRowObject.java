package com.doughtnerd.pod.excel.abstracts;

import com.doughtnerd.pod.excel.ExcelCellObject;
import com.doughtnerd.pod.excel.enums.ExcelFormatType;

/**
 * This is an abstract class that represents an object that is writable as a row
 * in an excel document by use of the toCellObjectArray method. It also provides
 * the createCell method that allows for the quick creation of ExcelCellObjects
 * when filling in the toCellObjectArray method.
 * 
 * @see #toCellObjectArray()
 * @see #createCell(Object)
 * @author Christopher Carlson
 *
 */
public abstract class ExcelRowObject {

	/**
	 * Converts this excel object into an array that is writable to an excel
	 * file as a complete row.
	 * 
	 * @return An ExcelCellObject[] of all of this ExcelRowObject's desired
	 *         field values.
	 */
	public abstract ExcelCellObject[] toCellObjectArray();

	/**
	 * This method provides a quick way for the user to wrap an ExcelRowObject
	 * field into an ExcelCellObject - which contains formatting data, allowing faster creation of
	 * ExcelCellObject arrays.
	 * 
	 * @see #toCellObjectArray()
	 * @param value
	 *            The Object that will be set the new ExcelCellObject's value.
	 * @return The ExcelCellObject that is created by invoking this method.
	 */
	protected ExcelCellObject createCell(Object value) {
		return new ExcelCellObject(value);
	}
	
	protected ExcelCellObject createCell(Object value, ExcelFormatType format){
		return new ExcelCellObject(value, format);
	}
}
