package com.doughtnerd.pod.excel.unitTests;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.junit.Assert;
import org.junit.Test;

import com.doughtnerd.pod.excel.ExcelCellObject;
import com.doughtnerd.pod.excel.abstracts.ExcelReader;
import com.doughtnerd.pod.excel.abstracts.ExcelRowObject;

public class ExcelReaderTests {

	@Test
	public void sheetNotFoundTest() {

	}

	class TestReader extends ExcelReader<TestData>{

		public TestReader(File file) throws IOException {
			super(file);
		}

		@Override
		protected TestData extractItem(Row row) {
			Iterator<Cell> iter = row.cellIterator();
			TestData data = new TestData();
			while(iter.hasNext()){
				data.data.add(iter.next());
			}
			return data;
		}
	}
	
	class TestData extends ExcelRowObject{
		
		public ArrayList<Object> data;
		
		public TestData(){
			this.data = new ArrayList<Object>();
		}

		@Override
		public ExcelCellObject[] toCellObjectArray() {
			ArrayList<ExcelCellObject> list = new ArrayList<>();
			for(Object o : data){
				list.add(this.createCell(o));
			}
			return list.toArray(new ExcelCellObject[0]);
		}
	}
}
