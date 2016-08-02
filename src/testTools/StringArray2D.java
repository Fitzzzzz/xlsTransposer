package testTools;

public class StringArray2D {

	private String[][] tab;
	
	public String[][] getTab() {
		return tab;
	}

	public void setTab(String[][] tab) {
		this.tab = tab;
	}

	public StringArray2D(String[][] tab) {
		this.tab = tab;
	}

	

	public void print() {
		StringArray line;
		for (String[] t : tab) {
			line = new StringArray(t);
			line.print();
		}
	}
	
	
}
