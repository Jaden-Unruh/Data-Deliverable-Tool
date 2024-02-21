package dataDeliverableTool;

import java.util.MissingResourceException;
import java.util.ResourceBundle;

/**
 * Class to fetch all externalized strings
 * @author Jaden Unruh
 * @see Main
 */
public class Messages {
	private static final ResourceBundle RESOURCE_BUNDLE = ResourceBundle.getBundle("dataDeliverableTool/messages"); //$NON-NLS-1$

	/**
	 * Returns the String with the associated key in the messages.properties file
	 * @param key the key to use
	 * @return the corresponding String
	 */
	public static String getString(String key) {
		try {
			return RESOURCE_BUNDLE.getString(key);
		} catch (MissingResourceException e) {
			return '!' + key + '!';
		}
	}
}
