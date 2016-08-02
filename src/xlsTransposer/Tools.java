package xlsTransposer;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;


/**
 * Different tools for reading and writing from a sheet to another
 * @author hamme
 *
 */

public class Tools {

	/**
	 * Constructor
	 * @param input
	 * 		sheet that will be read
	 * @param output
	 * 		sheet in which it will be written
	 */
	public Tools(HSSFSheet input, XSSFSheet output) {
		this.input = input;
		this.output = output;
	}
	/**
	 * sheet that will be read
	 */
	private HSSFSheet input;
	/**
	 * sheet in which it will be written
	 */
	private XSSFSheet output;
	/**
	 * Last column of the input sheet
	 */
	private int lastColumn;
	
	public int getLastColumn() {
		return lastColumn;
	}
	public void setLastColumn(int columnLimit) {
		this.lastColumn = columnLimit;
	}
	
	/**
	 * Copies from {@link Tools#input} (starting at line inputStart) into {@link Tools#output} 
	 * (starting at line outputStart) a number of lines equal to length.
	 * @param inputStart
	 * 		the sheet to copy from
	 * @param outputStart
	 * 		the sheet to copy in
	 * @param length
	 * 		the number of lines to copy
	 */
	public void copy(int inputStart, int outputStart, int length) {

		HSSFRow iRow;
		XSSFRow oRow;
		for (int i = 0; i < length; i++) {
			
			iRow = input.getRow(inputStart + i);
			oRow = output.createRow(outputStart + i);
			
			// Checking if the row isn't null.
			if (iRow != null) {

				int eol = iRow.getLastCellNum();
				
				// Copying each cell of the row
				for (int j = 0; j < eol; j++) {
					
					Cell iCell = iRow.getCell(j, Row.CREATE_NULL_AS_BLANK);
					Cell oCell = oRow.createCell(j);
					
					switch (iCell.getCellType()) {
					case Cell.CELL_TYPE_NUMERIC:
						// If the format is a date, we have to copy it as a date
					 	if (DateUtil.isCellDateFormatted(iCell)) {
					 		oCell.setCellValue(iCell.getDateCellValue());
 					 	}
					 	else {
					 		oCell.setCellValue(iCell.getNumericCellValue());
					 	}	
						break;
					case Cell.CELL_TYPE_STRING:
						oCell.setCellValue(iCell.getStringCellValue());
						break;
					}

				}

			}

		}
		
	}
	
	/**
	 * Extracts a line from the {@link Tools#input}.
	 * @param rowId
	 * 		The number of the row.
	 * @return
	 * 		The row as a cell array
	 */
	public Cell[] extractLine(int rowId) { 
		
		HSSFRow row = input.getRow(rowId);
		Cell[] line = new Cell[lastColumn + 1];		
		
		for (int i = 0; i <= lastColumn; i++) {			
			line[i] = row.getCell(i, Row.CREATE_NULL_AS_BLANK);
		}
		return line;
	}
	
	/**
	 * Extracts certain following cells from a row from the {@link Tools#input}.
	 * @param rowId
	 * 		The number of the row to extract cells from
	 * @param start
	 * 		The cell to start with
	 * @param end
	 * 		The last cell to extract
	 * @return
	 * 		The extracted cells in the form of an array of cells
	 */
	public Cell[] extractLine(int rowId, int start, int end) { 
		
		HSSFRow row = input.getRow(rowId);
		Cell[] line = new Cell[end - start + 1];
		
		
		for (int i = start; i <= end; i++) {
			
			line[i - start] = row.getCell(i, Row.CREATE_NULL_AS_BLANK);
		}
		return line;
	}
	/**
	 * Extracts certain following cells from a row from the {@link Tools#input}. 
	 * Extracts Null as Blank.
	 * Fills the comment array in parameter with the found commentaries.
	 * @param rowId
	 * 		The number of the row to extract cells from
	 * @param start
	 * 		The cell to start with
	 * @param end
	 * 		The last cell to extract
	 * @param
	 * 		The Comment array to fill with the commentaries
	 * @return
	 * 		The extracted cells in the form of an array of cells
	 */
	public Cell[] extractLine(int rowId, int start, int end, Comment[] comm) { 
		
		HSSFRow row = input.getRow(rowId);
		Cell[] line = new Cell[end - start + 1];

		for (int i = start; i <= end; i++) {			
			line[i - start] = row.getCell(i, Row.CREATE_NULL_AS_BLANK);
			Comment comment = row.getCell(i, Row.CREATE_NULL_AS_BLANK).getCellComment();
			if (comment != null) {
				System.out.println("comm non nul");
				comm[i - start] = comment;
			 }
		}
		return line;
	}
	/**
	 * Write a line in the {@link Tools#output} composed of arrays of Cell and String
	 * @param rowId
	 * 		The number of the row to write into
	 * @param beginning
	 * 		The cell array the line will start with
	 * @param end
	 * 		The string array the line will end with (one cell per String)
	 */
	public void writeLine(int rowId, Cell[] beginning, String[] end) {
		
		XSSFRow row = output.createRow(rowId);
		
		// Writing the beginning cell array
		for (int i = 0; i < beginning.length; i++) {
			Cell cell = row.createCell(i);
			switch (beginning[i].getCellType()) {
			case Cell.CELL_TYPE_NUMERIC:
				if (DateUtil.isCellDateFormatted(beginning[i])) {
					cell.setCellValue(beginning[i].getDateCellValue());
				}
				else {
					cell.setCellValue(beginning[i].getNumericCellValue());
				}
				break;
			case Cell.CELL_TYPE_STRING:
				cell.setCellValue(beginning[i].getStringCellValue());
				break;
			}
			
		}
		// Writing the end String array
		for (int i = 0; i < end.length; i++) {
			Cell cell = row.createCell(beginning.length + i);
			cell.setCellValue(end[i]);
		}
		
	}
	/**
	 * Writes a line in {@link Tools#output}  composed of arrays of Cell and String
	 * @param rowId
	 * 		The number of the row to write into
	 * @param beginning
	 * 		First array to be copied in
	 * @param middle
	 * 		Second array to be copied in
	 * @param end
	 * 		Last Array to be copied in
	 */
	public void writeLine(int rowId, Cell[] beginning, String[] middle, Cell[] end) {
		
		XSSFRow row = output.createRow(rowId);
		
		// Writing beginning
		for (int i = 0; i < beginning.length; i++) {
			Cell cell = row.createCell(i);
			switch (beginning[i].getCellType()) {
			case Cell.CELL_TYPE_NUMERIC:
				if (DateUtil.isCellDateFormatted(beginning[i])) {
					cell.setCellValue(beginning[i].getDateCellValue());
//					System.out.println(" FOUND A DATE " + beginning[i].getRowIndex() + ":" + beginning[i].getColumnIndex()); // TODO : TBR
				}
				else {
					cell.setCellValue(beginning[i].getNumericCellValue());
				}
				break;
			case Cell.CELL_TYPE_STRING:
				cell.setCellValue(beginning[i].getStringCellValue());
				break;
			}
			
		}
		// Writing middle
		for (int i = 0; i < middle.length; i++) {
			Cell cell = row.createCell(beginning.length + i);
			cell.setCellValue(middle[i]);
		}
		// Writing end
		for (int i = 0; i < end.length; i++) {
			Cell cell = row.createCell(beginning.length + middle.length+ i);
			switch (end[i].getCellType()) {
			case Cell.CELL_TYPE_NUMERIC:
				if (DateUtil.isCellDateFormatted(end[i])) {
					cell.setCellValue(end[i].getDateCellValue());
				}
				else {
					cell.setCellValue(end[i].getNumericCellValue());
				}
				break;
			case Cell.CELL_TYPE_STRING:
				cell.setCellValue(end[i].getStringCellValue());
				break;
			}
		}
		
	}
	/**
	 * Writes a line in {@link Tools#output}  composed of arrays of Cell and String
	 * @param rowId
	 * 		The number of the row write in.
	 * @param beginning
	 * 		First array to write of
	 * @param second
	 * 		Second array to write of
	 * @param third
	 * 		Third array to write of
	 * @param end
	 * 		Last array to write of
	 */
	public void writeLine(int rowId, Cell[] beginning, String[] second, Cell[] third, String[] end) {
		
		XSSFRow row = output.createRow(rowId);
		// Writing beginning
		for (int i = 0; i < beginning.length; i++) {
			Cell cell = row.createCell(i);
			switch (beginning[i].getCellType()) {
			case Cell.CELL_TYPE_NUMERIC:
				if (DateUtil.isCellDateFormatted(beginning[i])) {
					cell.setCellValue(beginning[i].getDateCellValue());
				}
				else {
					cell.setCellValue(beginning[i].getNumericCellValue());
				}
				break;
			case Cell.CELL_TYPE_STRING:
				cell.setCellValue(beginning[i].getStringCellValue());
				break;
			}
			
		}		
		// Writing second
		for (int i = 0; i < second.length; i++) {
			Cell cell = row.createCell(beginning.length + i);
			cell.setCellValue(second[i]);
		}		
		// Writing third
		for (int i = 0; i < third.length; i++) {
			Cell cell = row.createCell(beginning.length + second.length+ i);
			switch (third[i].getCellType()) {
			case Cell.CELL_TYPE_NUMERIC:
				if (DateUtil.isCellDateFormatted(third[i])) {
					cell.setCellValue(third[i].getDateCellValue());
				}
				else {
					cell.setCellValue(third[i].getNumericCellValue());
				}
				break;
			case Cell.CELL_TYPE_STRING:
				cell.setCellValue(third[i].getStringCellValue());
				break;
			}
		}
		// Writing end
		for (int i = 0; i < end.length; i++) {
			Cell cell = row.createCell(beginning.length + second.length + third.length + i);
			cell.setCellValue(end[i]);
		}
		
	}
	/**
	 * Writes an array of Cell into a created row
	 * @param rowId
	 * 		The number of the row to write into
	 * @param line
	 * 		The array of cell to write
	 */
	public void writeLine(int rowId, Cell[] line) {
		
		XSSFRow row = output.createRow(rowId);
		for (int i = 0; i < line.length; i++) {
//			System.out.println("writeLine de la case " + i);
			Cell cell = row.createCell(i);
			switch (line[i].getCellType()) {
			case Cell.CELL_TYPE_NUMERIC:
				if (DateUtil.isCellDateFormatted(line[i])) {
//					System.out.println(" FOUND A DATE " + line[i].getRowIndex() + ":" + line[i].getColumnIndex()); // TODO : TBR
					
					cell.setCellValue(line[i].getDateCellValue());
				}
				else {
					cell.setCellValue(line[i].getNumericCellValue());
				}
				break;
			case Cell.CELL_TYPE_STRING:
				cell.setCellValue(line[i].getStringCellValue());
				break;
			}
			
		}
		
	}
	/**
	 * Writes a line in {@link Tools#output}  from arrays of Cell, int and Cells
	 * @param rowId
	 * 		The line to write at
	 * @param first
	 * 		First array to write of
	 * @param second
	 * 		Second part to write of
	 * @param third
	 * 		Third part to write of
	 * @param forth
	 * 		Forth part to write of
	 * @param last
	 * 		Last part to write of
	 */
	public void writeline(int rowId, Cell[] first, int second, int third, Cell forth, Cell[] last) {

		
		XSSFRow row = output.createRow(rowId);
		// Writing first
		for (int i = 0; i < first.length; i++) {
			Cell cell = row.createCell(i);
			switch (first[i].getCellType()) {
			case Cell.CELL_TYPE_NUMERIC:
				if (DateUtil.isCellDateFormatted(first[i])) {
					cell.setCellValue(first[i].getDateCellValue());
				}
				else {
					cell.setCellValue(first[i].getNumericCellValue());
				}
				break;
			case Cell.CELL_TYPE_STRING:
				cell.setCellValue(first[i].getStringCellValue());
				break;
			}
			
		}

		// Writing second
		Cell cell = row.createCell(first.length);
		cell.setCellValue(second);
		
		// Writing third
		cell = row.createCell(first.length + 1);
		cell.setCellValue(third);
		
		// Writing forth
		cell = row.createCell(first.length + 2);
		switch (forth.getCellType()) {
		case Cell.CELL_TYPE_NUMERIC:
			if (DateUtil.isCellDateFormatted(forth)) {
				cell.setCellValue(forth.getDateCellValue());
			}
			else {
				cell.setCellValue(forth.getNumericCellValue());
			}
			break;
		case Cell.CELL_TYPE_STRING:
			cell.setCellValue(forth.getStringCellValue());
			break;
		}
		
		// Writing last
		for (int i = 0; i < last.length; i++) {
			cell = row.createCell(first.length + 3 + i);
			switch (last[i].getCellType()) {
			case Cell.CELL_TYPE_NUMERIC:
				if (DateUtil.isCellDateFormatted(last[i])) {
					cell.setCellValue(last[i].getDateCellValue());
				}
				else {
					cell.setCellValue(last[i].getNumericCellValue());
				}
				break;
			case Cell.CELL_TYPE_STRING:
				cell.setCellValue(last[i].getStringCellValue());
				break;
			}
		}
		
	}
	
	/**
	 * Writes a message in a cell in the output sheet.
	 * @param rowId
	 * 		The row of the cell to write in
	 * @param columnId
	 * 		The column of the cell to write in
	 * @param msg
	 * 		The message to write in the cell
	 */
	public void writeCell(int rowId, int columnId, String msg) {
		
		XSSFRow row = output.getRow(rowId);
		Cell cell = row.createCell(columnId);
		cell.setCellValue(msg);
		
	}
	/**
	 * Checks if a certain row is at or beyond EOF in the input sheet. We consider here EOF 
	 * as being at the first empty or null row encountered.
	 * @param j
	 * 		The index of the row to check.
	 * @see Tools#isRowEmpty(XSSFRow)
	 * @return
	 * 		True if EOF reached, false otherwise.
	 */
	public boolean isItEOF(int j) {
		HSSFRow row = input.getRow(j);
		if (row == null) {
			return true;
		}
		// If the row is empty
		if (isRowEmpty(row)) {
			return true;
		}
		else {
			return false;
		}
	}
	/**
	 * Checks if a row is empty.
	 * @param row
	 * 		The row number
	 * @return
	 * 		true if empty, false otherwise
	 */
	public static boolean isRowEmpty(HSSFRow row) {
	    for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
	        Cell cell = row.getCell(c);
	        if (cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK)
	            return false;
	    }
	    return true;
	}
	/**
	 * Checks if a column is empty starting a certain line in the output sheet.
	 * @param column
	 * 		The column to check
	 * @param firstRow
	 * 		The row to start with
	 * @return
	 * 		true if empty, false otherwise
	 */
	public boolean isColumnEmpty(int column, int firstRow) {
		

		HSSFRow row = input.getRow(firstRow);
		
		while (row != null) {
			Cell c = row.getCell(column, Row.RETURN_BLANK_AS_NULL);
			if (c != null) {
				return false;
			}
			row = input.getRow(firstRow++);
		}
		return true;
		
	}
	
	/**
	 * Checks if a column is empty starting a certain line in a certain sheet.
	 * @param column
	 * 		The column to check
	 * @param firstRow
	 * 		The row to start with
	 * @param sheet
	 * 		The sheet to check
	 * @return
	 * 		True if empty, False otherwise
	 */
	public static boolean isColumnEmpty(int column, int firstRow, XSSFSheet sheet) {
		

		XSSFRow row = sheet.getRow(firstRow);
		
		while (row != null) {
			Cell c = row.getCell(column, Row.RETURN_BLANK_AS_NULL);
			if (c != null) {
				return false;
			}
			row = sheet.getRow(firstRow++);
		}
		return true;
		
	}
	 
	/**
	 * Divides an array of cells between two other arrays. It fills the first completely first and 
	 * then tries filling the second one.
	 * The sum of the lengths of b and c must be superior or equal to the length of a.
	 * @param a
	 * 		The array to divide
	 * @param b
	 * 		The array receiving the first part of the array to divide.
	 * @param c
	 * 		The array receiving the second part of the array to divide.
	 */
	public static void divide(Cell[] a, Cell[] b, Cell[] c) {
		
		int bLength = b.length;
		for (int i = 0; i < bLength; i++) {
			b[i] = a[i];
		}
		for (int i = 0; i < c.length; i++) {
			c[i] = a[i + bLength];
		}
	}
	/**
	 * Regroup tow arrays into one. It copies the first completely first.	
	 * The sum of the lengths of b and c must be inferior or equal to the length of a.
	 * @param a
	 * 			The array which will receive the other two.
	 * @param b
	 * 			The first array to transfer cells from.
	 * @param c
	 * 			The second array whose cells are going to be copied.
	 */
	public static void regroup(Cell[] a, Cell[] b, Cell[] c) {
		
		int bLength = b.length;
		for (int i = 0; i < bLength; i++) {
			a[i] = b[i];
		}
		for (int i = 0; i < c.length; i++) {
			a[i + bLength] = c[i];
		}
	}
	/**
	 * Fills an array of Cell with another smaller array of Cell starting at a certain 
	 * index in the receiving array.
	 * a.length() - start must be equal or superior to b.length
	 * @param a
	 * 		Array that will be filled.
	 * @param b
	 * 		Array whose cells are extracted
	 * @param start
	 * 		Index to start from in a.
	 */
	public static void fill(Cell[] a, Cell[] b, int start) {
		
		for (int i = 0; i < b.length; i++) {
			a[start + i] = b[i];
			
		}
	}
	/**
	 * Prints the content of a cell
	 * @param c
	 * 		The cell to be printed.
	 */
	public static void printCell(Cell c) {
		
		if (c == null) {
			System.out.print("NULL #");
		}
		else {
			switch (c.getCellType()) {
			case Cell.CELL_TYPE_NUMERIC:
				if (DateUtil.isCellDateFormatted(c)) {
					System.out.print(c.getDateCellValue());
				}
				else {
					System.out.print(c.getNumericCellValue());
				}
				break;
			case Cell.CELL_TYPE_STRING:
				System.out.print(c.getStringCellValue());
				break;
			}
		}
	}
	
	
	
	
	
	
	
}
