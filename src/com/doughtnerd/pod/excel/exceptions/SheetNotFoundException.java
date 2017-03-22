package com.doughtnerd.pod.excel.exceptions;

/**
 * This class represents an exception that can be thrown if a sheet resulting
 * from a getSheetAt or an iteration through a workbooks sheets returns null.
 * 
 * @author Christopher Carlson
 *
 */
public final class SheetNotFoundException extends Exception {

	/**
	 * The serial ID.
	 */
	private static final long serialVersionUID = -5638479421197209450L;
	/**
	 * Exception message.
	 */
	private String message;

	/**
	 * Creates a new SheetNotFoundException with a null message.
	 */
	public SheetNotFoundException() {
		this(null);
	}

	/**
	 * Creates a new SheetNotFoundException with the specified message.
	 */
	public SheetNotFoundException(String message) {
		this.message = message;
	}

	/**
	 * Returns this SheetNotFoundException object's message
	 */
	public String getMessage() {
		return this.message;
	}
}
