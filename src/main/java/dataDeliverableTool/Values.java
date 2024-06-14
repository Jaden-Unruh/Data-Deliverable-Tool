package dataDeliverableTool;

import java.util.ResourceBundle;

/**
 * Class to fetch all externalized values
 * @author Jaden
 * @see Main
 *
 */
public class Values {
	private static final ResourceBundle RESOURCE_BUNDLE = ResourceBundle.getBundle("dataDeliverableTool/values");
	
	/**
	 * Returns the integer with the associated key in the values.properties file
	 * @param key the key to use
	 * @return the corresponding int
	 */
	public static int getValue(String key) {
		return  Integer.parseInt(RESOURCE_BUNDLE.getString(key));
	}
}
