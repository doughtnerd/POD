package com.doughtnerd.pod.excel.abstracts;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Row;

import com.doughtnerd.pod.excel.abstracts.ExcelReader;

/**
 * This class is an alternate implementation of ExcelReader that allows for a
 * configuration file to be read in for configuring the reader.
 * 
 * @author Christopher Carlson
 *
 * @param <T>
 *            The data type this reader will be extracting from the Excel
 *            document.
 */
public abstract class ConfiguredExcelReader<T> {

	/**
	 * The properties this reader will use during its operations.
	 */
	private Properties props;

	/**
	 * The underlying ExcelReader object that will read the data.
	 */
	private DataReader reader;

	/**
	 * Creates a new ConfiguredExcelReader object.
	 * 
	 * @param configFile
	 *            The file used to configure this reader.
	 * @param dataFile
	 *            The data file this reader will operate on.
	 */
	public ConfiguredExcelReader(File configFile, File dataFile) {
		Properties prop = new Properties();
		InputStream input = null;
		try {
			reader = new DataReader(dataFile);
			input = new FileInputStream(configFile);
			prop.load(input);
			props = prop;
			configureReader(this.props);
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (input != null) {
				try {
					input.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}

	/**
	 * Called by the constructor to initialize any variables the reader might
	 * contain or need to use before the extraction process begins.
	 * 
	 * @param properties
	 *            The properties read in from the configuration file upon
	 *            construction of this reader.
	 */
	public abstract void configureReader(Properties properties);

	/**
	 * How this reader will extract data from a given row in the excel document.
	 * 
	 * @param row
	 *            The current row this reader is operating on.
	 * @param properties
	 *            The properties read in from the configuration file.
	 * @return The data extracted from the row.
	 */
	public abstract T extractItem(Row row, Properties properties);

	/**
	 * @return The underlying data reader.
	 */
	public DataReader getReader() {
		return reader;
	}

	public class DataReader extends ExcelReader<T> {

		public DataReader(File file) throws IOException {
			super(file);
		}

		@Override
		protected T extractItem(Row row) {
			return ConfiguredExcelReader.this.extractItem(row, props);
		}
	}
}
