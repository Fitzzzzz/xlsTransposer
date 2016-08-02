package xlsTransposer;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/**
 * A class describing a sheet. Can be a sheet in which it will be written (OutputFile) OR read (Inputfile) but not both.
 * 
 * @author hamme
 *
 */

public class InOutFile {


	/**
	 * The line in which it will be written or read next. Starts at 0.
	 */
	private int currentLine; 
	

	public int getCurrentLine() {
		return currentLine;
	}
	
	public void setCurrentLine(int l) {
		this.currentLine = l;
	}
	
	/**
	 * The sheet of the input file (with a .xls extension)
	 */
	private HSSFSheet hSheet;
	
	public HSSFSheet getHSheet() {
		return hSheet;
	}

	public void setHSheet(HSSFSheet sheet) {
		this.hSheet = sheet;
	}
	
	/**
	 * The sheet of the output file (with a .xlsx extension)
	 */
	private XSSFSheet xSheet;
	
	public XSSFSheet getXSheet() {
		return xSheet;
	}

	public void setXSheet(XSSFSheet sheet) {
		this.xSheet = sheet;
	}



	
	/**
	 * Constructer for handling the input file.
	 * Sets the the current line to 0.
	 * @param {@link InOutFile#hSheet}
	 * @see InOutFile#currentLine
	 */
	public InOutFile(HSSFSheet sheet) {
		this.currentLine = 0;
		this.hSheet = sheet;
	}
	/**
	 * Constructer for handling the output file.
	 * Sets the the current line to 0.
	 * @param {@link InOutFile#hSheet}
	 * @see InOutFile#currentLine
	 */
	public InOutFile(XSSFSheet sheet) {
		this.currentLine = 0;
		this.xSheet = sheet;
	}
	
	
	/**
	 * Increments the current line.
	 * @see InOutFile#currentLine
	 */
	public void incrLine() {

		currentLine++;
	}
}

