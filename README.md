# **POD**
POD (Poor Obfuscation Derivative) is an abstraction library that sits on top of Apache POI to make the reading/writing/editing of Microsoft Office documents more user friendly and intuitive. 

***Currently POD only handles Excel documents***

## **About**

Apache POI is an extensive library designed by the Apache foundation for handling the read/write/edit of Microsoft Office OLE2 format documents.

*If you wish to learn more about Apache POI or review its documentation, follow this [link](https://poi.apache.org/).*

However, because the POI library is so large, has so many deprecated uses, and is constantly being improved, using it can be a dismal and dizzying experience that can discourage many new users of the library. That is where **POD** comes in.

POD contains (as of version 1.0.0) only 6 main classes and 1 enum that can handle the reading, writing, and editing of MS Office documents *(Only Excel is available as of 1.0.0)*. As compared to the 2240 total Classes in the Apache POI library. 

Though the POD library is completely stable and will likely satisfy 90% of the average user's needs, it is missing abstractions for a lot of the more advanced features that the Apache POI library is capable of. When such situations occur, users are encouraged to either create the abstraction themselves then notify the developer of POD so the implementation can be added to the library, attempt to use the Apache POI API to fill in gaps as more of a '*patchwork*' solution, or wait for future versions of POD.

**This software is in no way endorsed by or affiliated with the Apache Foundation, its members, etc.** However, this software does use the Apache POI library under the hood.

## **Usage**

### Reading Excel Files
Reading in Excel files is a rather simple process.

Here is a short example class designed to read Account objects from an excel file.
	
	public class AccountReader extends ExcelReader<Account>{
    
    	public AccountReader(File file) throws IOException{
        	super(file);
        }
    
    	@Override
        protected Account extractItem(Row row){
        	Account account = new Account();
            account.setBranchCode(row.getCell(1).toString());
            account.setTotalRevenue(row.getCell(17).getNumericCellValue());
            return account;
        }
	}
    
Now let's break down what is happening in this class.

1. We create a new Class that extends ExcelReader\<T> where T is the type of object that will be extracted from the Excel document (a row in excel represents a single instance of the aforementioned object). ExcelReader is POD's main abstract class for reading in Excel documents and extracting the rows into individual objects.
2. We implement the super constructor that takes a file parameter which represents the excel file we want to pull data from.
3. We implement the abstract method **extractItem(Row row)** from the ExcelReader class.
4. Within the **extractItem(Row row)** method, we create an instance of the object we want extracted, then we set object parameters to data from corresponding cells within the row that is passed into this method.
	* **row is a parameter that the user does not have to define, it is passed 		internally by other methods in the ExcelReader class.**
5. We return the Account object that we created.

*Note: As you can see, POD allows the user to have a fair amount of flexibility by allowing access to native Apache POI classes such as Row.*

Now that we have a class defining how our data is extracted from a file, let's put it to work.

Here is a short example class of how to use the ExcelReader object we created above.

	public class Example {
    	
        public static void main(String[] args){
        	File file = new File("C:/users/exampleUser/documents/example.xlsx");
        	AccountReader reader = new AccountReader(file);
            
            ArrayList<Account> list = reader.processSheet(0, true);
        }
    }
    
In the above class two main things are happening:

1. We create a new file pointing to our excel file and create a new AccountReader with that file.
2. We use ExcelReader's <code>public ArrayList\<T> processSheet(int sheetIndex, boolean headers)</code> method. This method will read in all T data that is contained in the given sheet, ignoring headers if the **headers** parameter is true (which tells the reader that yes, this sheet has headers, I don't want anything in the first row to count as data).

In addition to the **processSheet** method(s), there is also a <code>public TreeMap\<String, ArrayList<T>> processDocument(boolean headers)</code> method which will process all sheets in a document and return the data as a TreeMap.

*Read the Javadocs for a full listing of ExcelReader methods and their uses.*


***

### Writing Excel Files

POD contains an ExcelWriter class that is full of static methods that allow for various ways to write your data to a new Excel file.

Below is an example of one possible method of writing data to a new Excel file.

	public class ExampleWrite{
    	
        public static void main (String[] args){
        	File file = new File("C:/users/exampleUser/documents/example.xlsx");
        	AccountReader reader = new AccountReader(file);
            ArrayList<Account> list = reader.processSheet(0, true);
            
            for(Account a : list){
            	//Do some operation here.
            }
            
            Workbook newBook = ExcelWriter.writeNewSheetToNewWorkbook("xlsx", "MyNewSheet", null, list);
            
            ExcelWriter.writeWorkbookToFile(newBook, "C:/users/documents/newFile.xlsx");
        }
    }
    
In the above code two main things are happening:

1. We are taking our data that we performed operations on and writing it to a new sheet in a new workbook with the ExcelWriter method <code>public static \<T extends ExcelRowObject> Workbook writeNewSheetToNewWorkbook(String workbookType, String sheetName, List<String> headers, List\<T> data)</code>
2. We then save the Workbook object that is returned by that method to a file using thie <code>public static void writeWorkbookToFile(Workbook workbook, String path)</code>

Take a look at the method header in step 1 above. Well that's interesting isn't it? "What's an ExcelRowObject??", you might ask. "Why can't I just cram some random list of data into the **writeNewSheetToWorkbook** method???
"
That last question is also the answer. There is no way to reliably know how a user would want all of an object's data turned into a row when it comes time to write the object into an Excel file. That is what the abstract ExcelRowObject class is for.

#### ExcelRowObject Class

All objects a user wishes to write to an Excel file **must** extend the ExcelRowObject class. This class contains the abstract method <code>public abstract ExcelCellObject[] toCellObjectArray();</code>. This method tells the ExcelWriter class how you want to split an object's data when it gets written to a row in an Excel file. The class also contains the method <code>protected ExcelCellObject createCell(Object value)</code> to help wrap your data in ExcelCellObjects -Which we'll get to below, so you can more easily create the ExcelCellObject[] that the **toCellObjectArray()** method must return.

Below is a short example class of the Account object we used in above sections which demonstrates how to implement an object as an ExcelRowObject class

	public class Account extends ExcelRowObject{

        private String branchCode;
        private double totalRevenue;
        
        @Override
        public ExcelCellObject[] toCellObjectArray(){
        	ExcelCellObject[] arr = new ExcelCellObject[2];
            arr[0] = this.createCell(branchCode);
            arr[1] = this.createCell(totalRevenue);
            return arr;
        }
        
        //Getters, setters, etc below
    }
    
#### ExcelCellObject Class
Now as for what this ExcelCellObject class is, it's fairly simple - it represents a single cell in a row within an Excel file. This is an abstracted implementation of Apache POI's Cell class and all it does is contain data and format.

In essense, the ExcelCellObject class is simply a wrapper class that also allows the user to specify what type of format the data should have when it is written to an excel file such as:
* Wrap text - True or False
* Data format - General, US Currency, Percent

*Due to the changes that will be appearing in Apache POI when version 3.16 is complete, a definite and complete implementation of all formatting options available cannot be demonstrated. Please refer to the POD Javadocs for more info.
)
