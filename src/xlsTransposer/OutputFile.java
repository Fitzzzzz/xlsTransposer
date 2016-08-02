package xlsTransposer;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class OutputFile extends InOutFile {
	

	/**
	 * Constructor
	 * @param sheet 
	 * 		The sheet in which it will be written
	 */
	public OutputFile(XSSFSheet sheet) {
		super(sheet);
	}
	/**
	 * An array containing all the years (used in the case of a yearly period) 
	 * @see SheetCouple#isMonthly()
	 */
	private Cell[] years;
	
	
	public Cell[] getYears() {
		return years;
	}

	public void setYears(Cell[] years) {
		this.years = years;
	}
	/**
	 * An array containing all the years (used in the case of a monthly period)
	 * @see SheetCouple#isMonthly()
	 */
	private int[] yearsInt;
	
	public int[] getYearsInt() {
		return yearsInt;
	}

	public void setYearsInt(int[] yearsString) {
		this.yearsInt = yearsString;
	}
	/**
	 * An array containing all the months (used in the case of a monthly period)
	 * @see SheetCouple#isMonthly()
	 */
	public int[] months;
	
	public int[] getMonths() {
		return months;
	}

	public void setMonths(int[] months) {
		this.months = months;
	}
	/**
	 * The right part of the header that will be written in the sheet.
	 */
	private Cell[] rightHeader;
 
	public Cell[] getRightHeader() {
		return rightHeader;
	}

	public void setRightHeader(Cell[] rightHeader) {
		this.rightHeader = rightHeader;
	}
	/**
	 * The left part of the header that will be written in the sheet.
	 */
	private Cell[] leftHeader;


	public Cell[] getLeftHeader() {
		return leftHeader;
	}

	public void setLeftHeader(Cell[] leftHeader) {
		this.leftHeader = leftHeader;
	}
	/**
	 * The number of the first column of the comments in the output sheet.
	 */
	private int commentColumnId;
	

	public int getCommentColumnId() {
		return commentColumnId;
	}

	public void setCommentColumnId(int commentColumnId) {
		this.commentColumnId = commentColumnId;
	}
	/**
	 * All the values (one per period) of one line of the input sheet.
	 */
	private Cell[] values;
	
 	public Cell[] getValues() {
		return values;
	}
	public void setValues(Cell[] values) {
		this.values = values;
	}
	/**
	 * Divides the header of the input sheet between {@link OutputFile#leftHeader}, 
	 * {@link OutputFile#years} and {@link OutputFile#rightHeader}.
	 * @param header 
	 * 		The header of the input sheet
	 * @param serieNb
	 * 		The number of cells of the {@link OutputFile#leftHeader}
	 * @param firstRightHeader
	 * 		The index of the first cell of the {@link OutputFile#rightHeader} 
	 */
	public void divideHeader(Cell[] header, int serieNb, int firstRightHeader) {
		this.leftHeader = new Cell[serieNb];
		Cell[] left = new Cell[header.length - serieNb];
		Tools.divide(header, leftHeader, left);
		years = new Cell[firstRightHeader - serieNb];
		rightHeader = new Cell[left.length - years.length];
		Tools.divide(left, years, rightHeader);
	}
	
	
	/**
	 * Just the columns added in the output if yearly period.
	 * @see SheetCouple#isMonthly()
	 */
	public static final String[] periodValueYearly = {"period", "value"};
	/**
	 * Just the columns added in the output if monthly period.
	 * @see SheetCouple#isMonthly()
	 */
	public static final String[] periodValueMonthly = {"year", "month", "value"};
	/**
	 * Just the columns added for the comments.
	 */
	public static final String[] commentColumns = {"source", "comment", "statut"};
	
}

