package testTools;

public class StringArray {
	
	private String[] tab;
	
	public StringArray(String[] tab) {
		this.tab = tab;
	}

	public String[] getTab() {
		return tab;
	}

	public void setTab(String[] tab) {
		this.tab = tab;
	}

	@Override
	public String toString() {

		String line = "";
		for (String s : tab) {
			if (s != tab[0]) {
				line = line + " # " + s;
			}
			else
			{
				line = s;
			}
		}
		return line;
	}
	
	
	
	public void print() {
		System.out.println(toString());
	}


	
	
}
