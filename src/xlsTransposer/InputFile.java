package xlsTransposer;


import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;


public class InputFile extends InOutFile{


		/**
		 * Constructor
		 * @param sheet
		 * 		The sheet in which it will be read.
		 * @param linesToCopy
		 * 		Number of lines in the file that should be copy/pasted 
		 */
		public InputFile(HSSFSheet sheet, int linesToCopy) {
			super(sheet);
			this.linesToCopy = linesToCopy;
			allDone = false;
		}
		
		/**
		 * Constructor
		 * @param sheet
		 * 		The sheet in which it will be read.
		 */
		public InputFile(HSSFSheet sheet) {
			super(sheet);
			this.linesToCopy = 0;
			allDone = false;
		}

		/**
		 * A boolean that is true when the whole file has been read.
		 */
		private boolean allDone;
		
		public boolean isAllDone() {
			return allDone;
		}
		
		/**
		 * Number of lines in the file that should be copy/pasted 
		 */
		private int linesToCopy;

		public int getLinesToCopy() {
			return linesToCopy;
		}
		
		public void setLinesToCopy(int linesToCopy) {
			this.linesToCopy = linesToCopy;
		}

		/**
		 * The line number of the header 
		 * @see InputFile#header
		 */
		private int headerRow;
		
		public int getHeaderRow() {
			return headerRow;
		}
		
		public void setHeaderRow(int headerRow) {
			this.headerRow = headerRow;
		}
		
		/**
		 * Number of columns before the actual value or period.
		 */
		private int serieNb;
		
		public int getSerieNb() {
			return serieNb;
		}

		public void setSerieNb(int serieNb) {
			this.serieNb = serieNb;
		}

		/**
		 * An array that contains all cells composing the header
		 */
		private Cell[] header;

		public void setHeader(Cell[] header) {
			this.header = header;
		}

		public Cell[] getHeader() {
			return header;
		}

		/**
		 * The number (starting at column 0) of the last column of the file.
		 */
		private int lastColumn;
		
		public int getLastColumn() {
			return lastColumn;
		}

		public void setLastColumn(int columnLimit) {
			this.lastColumn = columnLimit;
		}

		/**
		 * The number (starting at column 0) of the last column containing values.
		 */
		public int lastPeriod;
		
		public int getLastPeriod() {
			return lastPeriod;
		}

		public void setLastPeriod(int lastPeriod) {
			this.lastPeriod = lastPeriod;
		}

		/**
		 * Will go thru the row of number rowId starting at column serieNb and search for the first blank or null cell. Its number will be stored in {@link InputFile#lastColumn}.
		 * Will also search in the row of number rowId + 1 starting from column serieNb for the first cell of type CELL_TYPE_STRING and store its number in {@link InputFile#lastPeriod}.
		 * @param rowId
		 * 		The Id (starting at 0) of the row in which it will be searched.
		 * @param serieNb
		 * 		The column of the cell to start the search with.
		 */
		public void findFirstBlankColumn(int rowId, int serieNb) {

			HSSFRow row = getHSheet().getRow(rowId);
			int eof = row.getLastCellNum();
			HSSFRow rowUnder = getHSheet().getRow(rowId + 1);
			boolean stringFound = false;

			for (int i = serieNb; i < eof; i++) {
				
				Cell cell = row.getCell(i, Row.RETURN_BLANK_AS_NULL);
				Cell cellUnder = rowUnder.getCell(i, Row.RETURN_BLANK_AS_NULL);
				
				// If the cell is empty or blank :
				if (cell == null) {
					
					// The last column was the one before.
					lastColumn = i - 1;
					return;	
				}
				
				// If the cell one row under is empty or blank :
				if (cellUnder != null) {
					// If the cell is of type String and one hadn't been found so far.
					if (!(stringFound) && (cellUnder.getCellType() == Cell.CELL_TYPE_STRING)) {
						// The last value was the one before.
						lastPeriod = i - 1;
						stringFound = true;
					}
				}
				
			}
			// In case we haven't found any blank or empty cell before EOF
			lastColumn = eof - 1;
			
		}
		
		
		
		
}
