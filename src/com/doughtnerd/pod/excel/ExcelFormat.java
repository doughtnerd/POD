package com.doughtnerd.pod.excel;

import com.doughtnerd.pod.excel.enums.ExcelFormatType;

/**
 * This class Handles setting cell formats to be added to a formatMap in an
 * excel object.
 * 
 * @author Christopher Carlson
 *
 */
public final class ExcelFormat {

	/**
	 * The format string that this format represents.
	 */
	private String formatString;

	/**
	 * Creates a new ExcelFormat object to help control excel cell styling.
	 * 
	 * @param type
	 *            An ExcelFormatType enum that represents the available format
	 *            types.
	 */
	public ExcelFormat(ExcelFormatType type) {
		switch (type) {
		case US_CURRENCY:
			this.formatString = "$#,#0.00";
			break;
		case PERCENT:
			this.formatString = "0.0%";
			break;
		case GENERAL:
			this.formatString = "General";
			break;
		default:
			throw new IllegalArgumentException("Invalid ExcelFormatType parameter");
		}
	}

	/**
	 * Returns this object's formatString.
	 * 
	 * @return A String representing the proper excel cell style for styling
	 *         purposes.
	 */
	public String getFormatString() {
		return formatString;
	}
}
