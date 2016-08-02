package testTools;

import org.apache.poi.ss.usermodel.Cell;

import xlsTransposer.Tools;

public class CellArray {

	private Cell[] tab;
	
	public CellArray(Cell[] tab) {
		this.tab = tab;
	}
	
	public String cellToString(Cell c) {
		
		String r = "";
		
		switch (c.getCellType()) {
		case Cell.CELL_TYPE_NUMERIC:
			r = String.valueOf(c.getNumericCellValue());
			break;
		case Cell.CELL_TYPE_STRING:
			r = c.getStringCellValue();
			break;
		}
		
		return r;
	}
	
	@Override
	public String toString() {
		
		String line = "";
		for (Cell c : tab) {
			if (c != tab[0]) {
				line = line + " # " + cellToString(c);
			}
			else
			{
				line = cellToString(c);
			}
		}
		
		return line;
	}
	
	public void print() {
		for (Cell c : tab) {
			Tools.printCell(c);
			System.out.println(" # ");
		}
		System.out.println("");
	}
}
