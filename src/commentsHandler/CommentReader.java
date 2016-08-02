package commentsHandler;

/**
 * A class handling a String : searching for the SOURCE, COMMENT & STATUT key-words in the String
 * and splitting the text following the key-words.  
 * @author hamme
 *
 */
public class CommentReader {


	/**
	 * Constructor. Constructs an array of String ({@link #words}) from a String by iteration. Each time
	 * a key-word is searched, the part of the array where it was detected is divided in three
	 * new Strings : the part before the key-word, the key-word and the part after.
	 * The position of the key-words in the array is stored in {@link CommentReader#position}.
	 * @param comment
	 * 		The string to split
	 * @param keyWords
	 * 		The key-words
	 */
	public CommentReader(String comment, String[] keyWords) {
		
		String[] tmp = {comment};
		for (String keyWord : keyWords) {
			tmp = split(tmp, keyWord);
		}
		words = new String[tmp.length];
		words = tmp;
		position = new int[keyWords.length];
		
		// Storing the positions of the key words in their order.
		for (int i = 0; i < keyWords.length; i++) {
			position[i] = getStringPosition(keyWords[i], words);
		}
	}
	
	/**
	 * An array containing the positions of the key-words in {@link CommentReader#words} in their order.
	 */
	private int position[];
	
	
	/**
	 * An array of Strings. Each String consists of either : some unimportant text standing alone
	 * (before the first key-word or the whole text if there was no key-word at all in the comment),
	 * a key-word (in which case, the following String is the text following the key-word) or 
	 * the text following a key-word as explained before.
	 */
	private String[] words;
	
	/**
	 * The position in the {@link CommentReader#words} array of the SOURCE: key-word.
	 */
	private int sourcePosition;
	
	public int getSourcePosition() {
		return sourcePosition;
	}

	/**
	 * The position in the {@link CommentReader#words} array of the COMMENT: key-word.
	 */
	private int commentPosition;
	
	public int getCommentPosition() {
		return commentPosition;
	}

	/**
	 * The position in the {@link CommentReader#words} array of the STATUT: key-word.
	 */
	private int statutPosition;
	
	public int getStatutPosition() {
		return statutPosition;
	}
	
	/**
	 * Searches from an array of String in each String if it contains a certain key-word. 
	 * If found, will split it between three Strings : the text before, the key-word, the text after.
	 * Will then create an array of the same size + 2 and copy the precedent inside plus the three new Strings
	 * at the place of the precedent place of the String where the key-word was found.
	 * @param text
	 * 		The array of String to search in
	 * @param keyWord
	 * 		The key-word searched for
	 * @return
	 * 		The same array if the key-word has not been found, the rebuilt one if found. 
	 */
	public String[] split(String[] text, String keyWord) {
		
		int tmp;
		String[] matchedArray;

		// We go thru the array
		for (int i = 0; i < text.length; i++) {
			
			// If the String contains the key-word
			if (text[i].contains(keyWord)) {
				
				tmp = i;
				matchedArray = text[i].split(keyWord);
				
				// The new array
				String[] newText = new String[text.length + 2];
				
				// Copy the Strings before the one in which it was found
				for (int j = 0; j < tmp; j++) {
					newText[j] = text[j];
				}
				
				// Insert the new Strings :
				// The text before the key-word
				newText[tmp] = matchedArray[0];
				// The key-word
				newText[tmp + 1] = keyWord;
				// The text after the key-word
				newText[tmp + 2] = matchedArray[1];
				
				// Copy the Strings after the one in which it was found
				for (int j = tmp + 1; j < text.length; j++) {
					newText[j + 2] = text[j];
				}
				
				// Return the new array
				return newText;
			}	
		}
		// Key-word not found, return the same array
		return text;	
	}
	/**
	 * Returns the position of a String in an Array if found.
	 * @param searchedWord
	 * 		The String we search for.
	 * @param text
	 * 		The array we search in.
	 * @return
	 * 		the position of the String if found, -1 if not found.
	 */
	private int getStringPosition(String searchedWord, String[] text) {
		return java.util.Arrays.asList(text).indexOf(searchedWord);
	}
	
	/**
	 * Finds a String in {@link CommentReader#words}
	 * @param keyWord
	 * 		The String to search for
	 * @return
	 * 		The place of the string in {@link CommentReader#words}, -1 if not found.
	 */
	public int getPosition(String keyWord) {
		return this.getStringPosition(keyWord, this.words);
	}
	// TODO : ADD exception if position == - 1
	/**
	 * Returns the text following a key-word in {@link CommentReader#words}
	 * @param keyWord
	 * 		The key-word to search for
	 * @return
	 * 		the text following a key-word in {@link CommentReader#words}
	 */
	public String getComment(String keyWord) {
		return words[getPosition(keyWord) + 1];
	}
}
