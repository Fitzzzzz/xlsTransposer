package xlsTransposer;

import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFComment;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import commentsHandler.CommentReader;

/**
 * Handles most of the job of the transposition. 
 * One is created for each input sheet + output sheet couple.
 * @author hamme
 *
 */
public class SheetCouple {

	/**
	 * Constructor, initializes an {@link #inputFile} and an {@link #outputFile}.
	 * @param iSheet
	 * 		The sheet in which it will be read.
	 * @param oSheet
	 * 		The sheet in which it will be written.
	 */
	public SheetCouple(HSSFSheet iSheet, XSSFSheet oSheet) {
		this.inputSheet = iSheet;
		this.outputSheet = oSheet;
		t = new Tools(inputSheet, outputSheet);
		this.inputCurrentLine = 0;
		this.outputCurrentLine = 0;
		this.inputFile = new InputFile(iSheet);
		this.outputFile = new OutputFile(oSheet);
		this.deletedValuesNb = 0;
	}
	/**
	 * Constructor, initializes an {@link #inputFile} and an {@link #outputFile}.
	 * @param iSheet
	 * 		The sheet in which it will be read.
	 * @param oSheet
	 * 		The sheet in which it will be written.
	 * @param linesToCopy
	 * 		The number of lines in the input sheet that should be copy/pasted.
	 */
	public SheetCouple(HSSFSheet iSheet, XSSFSheet oSheet, int linesToCopy) {
		this.inputSheet = iSheet;
		this.outputSheet = oSheet;
		t = new Tools(inputSheet, outputSheet);
		this.inputCurrentLine = 0;
		this.outputCurrentLine = 0;
		this.inputFile = new InputFile(iSheet, linesToCopy);
		this.outputFile = new OutputFile(oSheet);
	}
	
	/**
	 * The regular expression recognizing a monthly period.
	 */
	private static final String monthlyRegex = "[0-9]{4}_[0-9]{1,2}";
	/**
	 * The input file used to read.
	 */
	private InputFile inputFile;
	/**
	 * The outputFile used to write.
	 */
	public InputFile getInputFile() {
		return inputFile;
	}


	private OutputFile outputFile;
	
	public OutputFile getOutputFile() {
		return outputFile;
	}

	/**
	 * The sheet to be read.
	 */
	private HSSFSheet inputSheet;
	/**
	 * The sheet to be written in.
	 */
	private XSSFSheet outputSheet;
	/** 
	 * Tools helping the transposition.
	 */
	private Tools t;
	/**
	 * The current line of the input sheet.
	 */
	private int inputCurrentLine;
	/**
	 * The current line of the output sheet.
	 */
	private int outputCurrentLine;
	/** 
	 * Increments {@link SheetCouple#inputCurrentLine}
	 */
	public void incrInputLine() {
		this.inputCurrentLine++;
	}
	/** 
	 * Increments {@link SheetCouple#outputCurrentLine}
	 */
	public void incOutputLine() {
		this.outputCurrentLine++;
	}
	
	/**
	 * Number of lines (for the output) or cells (for the input) 
	 * that were skipped because they were empty. Is incremented dynamically in {@link #writeBody()}.
	 */
	private int deletedValuesNb;
	
	public int getDeletedValuesNb() {
		return deletedValuesNb;
	}
	
	private static final String[] keyWords = {"SOURCE:", "COMMENT:", "STATUT:"};
	
	/**
	 * Copies a number of succeeding lines (defined by length) from the input sheet starting at inputStart
	 * into the output sheet starting at outputStart.
	 * @param inputStart
	 * 		The index of the first line to copy.
	 * @param outputStart
	 * 		The index of the first line to copy in to.
	 * @param length
	 * 		The number of lines to copy.
	 */
	public void copy(int inputStart, int outputStart, int length) {
		
		t.copy(inputStart, outputStart, length);
		this.outputCurrentLine = this.outputCurrentLine + length;
		this.inputCurrentLine = this.inputCurrentLine + length;
	
	}
	/**
	 * @see InputFile#findFirstBlankColumn(int, int)
	 * @param rowId
	 * @param columnStart
	 */
	public void findFirstBlankColumn(int rowId, int columnStart) {
		inputFile.findFirstBlankColumn(rowId, columnStart);
		t.setLastColumn(inputFile.getLastColumn());
	}
	
	/**
	 * Extracts the header from the input sheet.
	 * @param headerRowId
	 */
	public void extractHeader(int headerRowId) {
		inputFile.setHeader(t.extractLine(headerRowId)); 
	}
	/**
	 * Divides the header from the input sheet between {@link OutputFile#leftHeader}, 
	 * {@link OutputFile#years} and {@link OutputFile#rightHeader}.
	 */
	public void divideHeader() {
		outputFile.divideHeader(inputFile.getHeader(), inputFile.getSerieNb(), inputFile.getLastPeriod() + 1);
	}
	/**
	 * Writes the header in the output sheet using {@link SheetCouple#findFirstBlankColumn(int, int)}, 
	 * {@link SheetCouple#extractHeader(int)} and {@link SheetCouple#divideHeader()}.
	 * Testing also if the period is yearly or monthly.
	 * @param headerRowId
	 * 		The index of the row containing the header in the input.
	 */
	public void writeHeader(int headerRowId) {
		
		findFirstBlankColumn(headerRowId, inputFile.getSerieNb());
		extractHeader(headerRowId);
		divideHeader();
		
		// If the period is yearly 
		if (!isMonthly()) {
			
			// Write the header lien
			t.writeLine(outputCurrentLine, outputFile.getLeftHeader(), OutputFile.periodValueYearly, outputFile.getRightHeader(), OutputFile.commentColumns);
			// Get the number of the first column of the comments
			this.outputFile.setCommentColumnId(outputFile.getLeftHeader().length + OutputFile.periodValueYearly.length + outputFile.getRightHeader().length);
		}
		// If the period is monthly
		else {
			int[] years = new int[outputFile.getYears().length];
			int[] months = new int[outputFile.getYears().length];
			for (int i = 0; i < years.length; i++) {
				String[] yearAndMonth = this.separeYearMonth(outputFile.getYears()[i]);
				// Retrieve the years
				years[i] = Integer.parseInt(yearAndMonth[0]);
				// Retrieve the months
				months[i] = Integer.parseInt(yearAndMonth[1]);
			}
			outputFile.setYearsInt(years);
			outputFile.setMonths(months);
			// Write the header
			t.writeLine(outputCurrentLine, outputFile.getLeftHeader(), OutputFile.periodValueMonthly, outputFile.getRightHeader(), OutputFile.commentColumns);
			// Set the number of the first comment column
			this.outputFile.setCommentColumnId(outputFile.getLeftHeader().length + OutputFile.periodValueMonthly.length + outputFile.getRightHeader().length);

		
		}
		
	}
	/**
	 * Write the body : For each line in the input, write one line in the output for each period.
	 * The first cells contain the same values for each line of one input line. 
	 * The following will contain the values, the period. At the end, the comments are displayed.
	 */
	public void writeBody() {
		
		int serieNb = inputFile.getSerieNb();
		int j = serieNb;
		boolean done = false;
		
		// If the period is yearly
		if (!isMonthly()) {
			System.out.println("non monthl");
			// While EOF of the input has not been reached
			while (!done) {
				
			    Comment[] comments = new Comment[inputFile.getLastPeriod() - serieNb + 1];
			    
			    outputFile.setLeftHeader(t.extractLine(j, 0, serieNb - 1));
			    outputFile.setValues(t.extractLine(j, serieNb, inputFile.getLastPeriod(), comments));
			    outputFile.setRightHeader(t.extractLine(j, inputFile.getLastPeriod() + 1, inputFile.getLastColumn()));
			    
			    Cell[] line = new Cell[serieNb + 2 + outputFile.getRightHeader().length];
			    
			    Tools.fill(line, outputFile.getLeftHeader(), 0);
			    Tools.fill(line, outputFile.getRightHeader(), serieNb + 2);
			    
			    // For one line of the input :
			    for (int i = 0; i < outputFile.getValues().length; i++) {
			    	
			    	if (!(outputFile.getValues()[i].getCellType() == Cell.CELL_TYPE_BLANK)) {
			    		
			    		// Get the year
				    	line[serieNb] = outputFile.getYears()[i];
				    	// Get the value
				    	line[serieNb + 1] = outputFile.getValues()[i];
				    	// Write the line
				    	t.writeLine(serieNb + i + (j - serieNb)*outputFile.getValues().length - deletedValuesNb, line);
			    	
				    	// If the comment isn't empty, we analyze it for keywords and write it if needed
				    	if (comments[i] != null) {				    		
				    		System.out.println("comment non nul");
				    		insertComment(comments[i], serieNb + i + (j - serieNb)*outputFile.getValues().length - deletedValuesNb);
				    	}
			    	}
			    	// If the cell is blank
			    	else {
			    		deletedValuesNb++;
			    	}
			    }
			    j++;
			    // Did we reach EOF?
			    done = t.isItEOF(j);
			    
		    }
		}
		// If the period is monthly
		else {

			// While EOF of the input has not been reached
			while (!done) {

			    Comment[] comments = new Comment[inputFile.getLastPeriod() - serieNb + 1];
			    outputFile.setLeftHeader(t.extractLine(j, 0, serieNb - 1));
			    outputFile.setValues(t.extractLine(j, serieNb, inputFile.getLastPeriod(), comments));
			    outputFile.setRightHeader(t.extractLine(j, inputFile.getLastPeriod() + 1, inputFile.getLastColumn()));

			    // For one line of the input :
			    for (int i = 0; i < outputFile.getValues().length; i++) {
			    	
			    	if (!(outputFile.getValues()[i].getCellType() == Cell.CELL_TYPE_BLANK)) {
			    		// Write the line
				    	t.writeline(
				    			serieNb + i + (j - serieNb)*outputFile.getValues().length - deletedValuesNb, 
				    			outputFile.getLeftHeader(), 
				    			outputFile.getYearsInt()[i], 
				    			outputFile.getMonths()[i], 
				    			outputFile.getValues()[i], 
				    			outputFile.getRightHeader()
				    			);
				    	// If the comment isn't empty, we analyze it for keywords and write it if needed
				    	if (comments[i] != null) {
				    		insertComment(comments[i],  serieNb + i + (j - serieNb)*outputFile.getValues().length - deletedValuesNb);
				    	}
			    	}
			    	// If the cell is blank
			    	else {
			    		deletedValuesNb++;
			    	}
			    }
			    j++;
			    // Did we reach the EOF?
			    done = t.isItEOF(j);	    
		    }
		}	
	}
	
	/**
	 * Goes thru a comment looking for the keywords and writes the portion following 
	 * a keyword in the output file at the given row.
	 * @param comment
	 * 		The comment to look thru
	 * @param rowId
	 * 		The row to write the comment in if needed
	 */
	public void insertComment(Comment comment, int rowId) {
		
		int commentIndex;
		// If the period is monthly 
		if (isMonthly()) {
			commentIndex = outputFile.getLeftHeader().length + OutputFile.periodValueMonthly.length + outputFile.getRightHeader().length;
		}
		// If the period is yearly
		else {
			commentIndex = outputFile.getLeftHeader().length + OutputFile.periodValueYearly.length + outputFile.getRightHeader().length;
		}
		
		CommentReader commentR = new CommentReader(comment.getString().getString(), keyWords);
		
		int outputRowId = rowId;
	     
		// If the source keyword has been detected.
		if (commentR.getPosition("SOURCE:") != -1) {
			System.out.println("Source found");
			t.writeCell(outputRowId, commentIndex, commentR.getComment("SOURCE:"));
		}
		// If the comment keyword has been detected
		if (commentR.getPosition("COMMENT:") != -1) {
			t.writeCell(outputRowId, commentIndex + 1, commentR.getComment("COMMENT:"));
			System.out.println("COMMENT found");
		}
		// If the statut keyword has been detected
		if (commentR.getPosition("STATUT:") != -1) {	
			t.writeCell(outputRowId, commentIndex + 2, commentR.getComment("STATUT:"));
			System.out.println("STATUT found");
		}
		
	}
	
	/**
	 * Will search for the commented cells. Once retrieved, the comments are analyzed and 
	 * key-words are searched for. The comments with key-words are conserved and written 
	 * in the respective columns. Most of the work is handled by CommentReader. 
	 */
	public void insertComments() {
		
		int valuesNumber = outputFile.getValues().length;
		int commentIndex;
		
		// If the period is monthly 
		if (isMonthly()) {
			commentIndex = outputFile.getLeftHeader().length + OutputFile.periodValueMonthly.length + outputFile.getRightHeader().length;
		}
		// If the period is yearly
		else {
			commentIndex = outputFile.getLeftHeader().length + OutputFile.periodValueYearly.length + outputFile.getRightHeader().length;
		}

		Map<CellAddress, HSSFComment> comments = inputSheet.getCellComments();
		    
		// For each comment
		for (java.util.Map.Entry<CellAddress, HSSFComment> e : comments.entrySet()) {
		      CellAddress loc = e.getKey();
		      Comment comment = e.getValue();		      
		      String[] keyWords = {"SOURCE:", "COMMENT:", "STATUT:"};
		      CommentReader commentR = new CommentReader(comment.getString().getString(), keyWords);

		      // Calculate the  in the output sheet of the cell whose comment is currently analyzed
		      int outputRowId = (loc.getRow() - inputFile.getLinesToCopy())*valuesNumber + inputFile.getLinesToCopy() + loc.getColumn() - inputFile.getSerieNb();
		     
		      // If the source keyword has been detected.
		      if (commentR.getPosition("SOURCE:") != -1) {
		    	  t.writeCell(outputRowId, commentIndex, commentR.getComment("SOURCE:"));
		      }
		      // If the comment keyword has been detected
				if (commentR.getPosition("COMMENT:") != -1) {
			      t.writeCell(outputRowId, commentIndex + 1, commentR.getComment("COMMENT:"));
		      }
		      // If the statut keyword has been detected
				if (commentR.getPosition("STATUT:") != -1) {
			      t.writeCell(outputRowId, commentIndex + 2, commentR.getComment("STATUT:"));
		      }
		}
	}

	/**
	 * Tests if the period is monthly or yearly by comparing the 
	 * {@link SheetCouple#monthlyRegex} to the first period cell.
	 * @return
	 * 		true if monthly, false is yearly
	 */
	public boolean isMonthly() {
		
		switch (outputFile.getYears()[0].getCellType()) {
		case Cell.CELL_TYPE_STRING: 
			return outputFile.getYears()[0].getStringCellValue().matches(monthlyRegex);
		default:
			return false;
		}
		
		
	}
	/**
	 * Separates the part concerning the year and the part concerning the month in a period cell
	 * @param yearWithMonthCell
	 * 		A cell matching the {@link SheetCouple#monthlyRegex} to be separated.
	 * @return.
	 * 		A String array containing the Strings corresponding to the year and to the month.
	 */
	private String[] separeYearMonth(Cell yearWithMonthCell) {
		String yearWithMonthString = yearWithMonthCell.getStringCellValue();
		String[] yearAndMonth = yearWithMonthString.split("_");
		if (yearAndMonth[1].matches("0[0-9]")) {
			yearAndMonth[1] = yearAndMonth[1].substring(1, 2);
		}
		return yearAndMonth;
	}
	
		
	
}
