package com.doughtnerd.pod.excel;

import java.util.Date;

import org.apache.poi.ss.usermodel.HorizontalAlignment;

import com.doughtnerd.pod.excel.enums.ExcelFormatType;

/**
 * This class represents an individual cell in Excel and is a wrapper class that
 * adds some functionality specific to writing excel documents. This class takes
 * an object as its value and allows the user to specify whether to wrap the
 * text in the cell this object is inserted into and also allows the user to
 * specify format.
 * 
 * @see #setDataFormat(ExcelFormat)
 * @see #setWrapText(boolean)
 * 
 * @author Christopher Carlson
 *
 */
public final class ExcelCellObject {

	/**
	 * The object this class is wrapped around, and the value that will be
	 * inserted into a cell.
	 */
	private Object value;

	/**
	 * The ExcelFormat that the cell containing this object's value will have
	 * applied.
	 */
	private ExcelFormat format;

	/**
	 * Whether or not the text in the cell that will contain this object's value
	 * should be wrapped.
	 */
	private boolean wrapText;

	/**
	 * Enum representing this object's vertical alignment property that will be
	 * added to the cell.
	 */
	private HorizontalAlignment hAlignment;

	/**
	 * <p>
	 * Creates a new ExcelCellObject. By default, when created, this object's
	 * format is set to GENERAL and {@link #wrapText} is set to false.
	 * </p>
	 * <p>
	 * If the value passed is not of type Character, String, Integer, Double,
	 * Date, or Boolean throws an IllegalArgumentException.
	 * </p>
	 * 
	 * @see ExcelFormatType
	 * @param value
	 *            The value this class is wrapped around.
	 */
	public ExcelCellObject(Object value) {
		this(value, null);
	}
	
	/**
	 * <p>
	 * Creates a new ExcelCellObject. By default, when created, this object's
	 * format is set to GENERAL and {@link #wrapText} is set to false.
	 * </p>
	 * <p>
	 * If the value passed is not of type Character, String, Integer, Double,
	 * Date, or Boolean throws an IllegalArgumentException.
	 * </p>
	 * 
	 * @see ExcelFormatType
	 * @param value
	 *            The value this class is wrapped around.
	 * @param format The format for the cell
	 *            
	 */
	public ExcelCellObject(Object value, ExcelFormatType format){
		if (!(value instanceof String) && !(value instanceof Integer) && !(value instanceof Double)
				&& !(value instanceof Character) && !(value instanceof Date) && !(value instanceof Boolean)) {
			throw new IllegalArgumentException(
					"Value passed to constructor must be of type Character, String, Integer, Double, Boolean, or Date. Value passed was: "
							+ value.getClass());
		}
		this.value = value;
		this.format = format!=null ? new ExcelFormat(format) : new ExcelFormat(ExcelFormatType.GENERAL);
		this.wrapText = false;
		this.hAlignment = HorizontalAlignment.LEFT;
	}

	/**
	 * Sets this object's data format to the new format.
	 * 
	 * @param format
	 *            New ExcelFormat to use in this object.
	 * @see ExcelFormat#ExcelFormat(ExcelFormatType)
	 */
	public void setDataFormat(ExcelFormat format) {
		this.format = format;
	}

	/**
	 * Get this object's data format.
	 * 
	 * @return This object's ExcelFormat representing the data format for the
	 *         object.
	 * @see ExcelFormat
	 */
	public ExcelFormat getDataFormat() {
		return this.format;
	}

	/**
	 * True if the text in the cell this object will represent should be
	 * wrapped. False otherwise.
	 * 
	 * @param wrapText
	 *            New value for wrapText.
	 */
	public void setWrapText(boolean wrapText) {
		this.wrapText = wrapText;
	}

	/**
	 * Whether or not the cell this object's value is added to should wrap its
	 * text.
	 * 
	 * @return This object's wrap text boolean.
	 * @see #wrapText
	 */
	public boolean getWrapText() {
		return this.wrapText;
	}

	/**
	 * Sets this object's horizontal alignment to the new hAligment enum value.
	 * 
	 * @param hAlignment
	 *            The new HorizontalAlignment enum value.
	 */
	public void setHorizontalAlignment(HorizontalAlignment hAlignment) {
		this.hAlignment = hAlignment;
	}

	/**
	 * Get this object's {@link HorizontalAlignment} enum value.
	 * 
	 * @return This objects horizontal alignment.
	 */
	public HorizontalAlignment getHorizontalAlignment() {
		return this.hAlignment;
	}

	/**
	 * This method returns the object that this class is wrapped around.
	 * 
	 * @return This object's wrapped value.
	 */
	public Object getValue() {
		return this.value;
	}
}
