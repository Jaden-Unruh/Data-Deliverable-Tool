package dataDeliverableTool;

import java.awt.Color;
import java.awt.Desktop;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.net.URL;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.concurrent.ExecutionException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.SwingWorker;
import javax.swing.WindowConstants;

import org.apache.commons.compress.utils.FileNameUtils;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Possible values for infoText element within the Window of Data Deliverable {@link Main#window}
 * @author Jaden Unruh
 * @since 0.0.1
 */
enum InfoText {
	ERROR, SELECT_PROMPT, DESKTOP, LOAD_SHEETS, INIT
}

/**
 * Primary class for Data Deliverable Tool.
 * @see <a href="https://github.com/Jaden-Unruh/data-deliverable-tool/">Data Deliverable Github</a>
 * @author Jaden Unruh
 * @since 0.0.1
 */
public class Main {

	/**
	 * Primary GUI window
	 */
	static JFrame window;

	/**
	 * Files currently selected - updates every time the user presses 'ok' within a file selection prompt
	 */
	static File[] selectedFiles = new File[2];

	/**
	 * Info bar within {@link #window}
	 * @see #infoText
	 */
	static JLabel info = new JLabel();

	/**
	 * Large objects representing the selected files - only update when the user selects 'run'
	 */
	static XSSFWorkbook deliverableBook, workbookBook;

	/**
	 * Current state of {@link #info}
	 * @see #getInfoText()
	 */
	static InfoText infoText;

	/**
	 * Writer for generated info text file - only defined in {@link #init()}
	 */
	static FileWriter writeToInfo;

	/**
	 * Entry method for {@link Main}
	 * @param args unused
	 */
	public static void main(String[] args) {
		openWindow();
	}

	/**
	 * Regex to pull old and new names from a line of <a href="file:/../resources/newNames.dat">newNames.dat</a>
	 */
	static final Pattern RENAME_LINE_PATTERN = Pattern.compile("(.+),(.+)");
	/**
	 * Initializes Data Deliverable tool - creates info file, pulls sheet renaming info
	 * @throws IOException if creation of info file (and/or writer) fails, or grabbing sheet renaming .dat file fails
	 */
	private static void init() throws IOException {
		updateInfo(InfoText.INIT);
		try (InputStream in = Main.class.getResourceAsStream("/NewNames.dat");
				BufferedReader reader = new BufferedReader(new InputStreamReader(in))) {
			String[] lines = reader.lines().toArray(String[]::new);
			for (String line : lines) {
				Matcher match = RENAME_LINE_PATTERN.matcher(line);
				if (match.find())
					nameMap.put(match.group(1), match.group(2));
			}
		} catch (IOException e) {
			e.printStackTrace();
		}

		String dateTime = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd_HH.mm.ss.SSS"));
		File infoFile = new File(
				String.format("%s\\Deliverable Program Information %s.txt", selectedFiles[0].getParent(), dateTime));
		infoFile.createNewFile();
		writeToInfo = new FileWriter(infoFile);
		writeToInfo.append(String.format(
				"Info log for deliverable tool run %s\nBelow will be any potential errors encountered while running the tool. If it's empty, no errors were reported.\n",
				dateTime));
	}
	
	/**
	 * Safely saves and closes info file writer and spreadsheet files
	 * @throws IOException if saving/closing fails
	 */
	private static void terminate() throws IOException {
		writeToInfo.close();
		FileOutputStream out = new FileOutputStream(selectedFiles[0]);
		deliverableBook.write(out);
		out.close();
		deliverableBook.close();
		workbookBook.close();
	}
	
	/**
	 * Opens and sets up the main window for the data deliverable tool
	 */
	private static void openWindow() {
		window = new JFrame("Data Deliverable Tool");
		window.setLayout(new GridBagLayout());
		window.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);

		window.add(new JLabel("Select a Deliverable CA file:"), simpleConstraints(0, 0, 2, 1));
		JButton selectDeliverable = new SelectButton(0);
		window.add(selectDeliverable, simpleConstraints(2, 0, 1, 1));

		window.add(new JLabel("Select a Workbook file:"), simpleConstraints(0, 1, 2, 1));
		JButton selectWorkbook = new SelectButton(1);
		window.add(selectWorkbook, simpleConstraints(2, 1, 1, 1));

		window.add(info, simpleConstraints(0, 2, 3, 1));

		JButton close = new JButton("Close");
		window.add(close, simpleConstraints(0, 3, 1, 1));

		JButton help = new JButton("Help");
		window.add(help, simpleConstraints(1, 3, 1, 1));

		final JButton run = new JButton("Run");
		window.add(run, simpleConstraints(2, 3, 1, 1));

		close.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				System.exit(0);
			}
		});

		help.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				String[] choices = { "Close", "Open GitHub" };
				if (JOptionPane.showOptionDialog(window,
						"Select two spreadsheets using the select buttons. They should both be of type .xlsx.\nThen, select `Run`, and the program will do the rest.\nFor more information, see the program GitHub readme.",
						"Help", JOptionPane.DEFAULT_OPTION, JOptionPane.INFORMATION_MESSAGE, null, choices,
						choices[0]) == 1) {
					if (Desktop.isDesktopSupported())
						try {
							Desktop.getDesktop()
									.browse(new URL("https://github.com/Jaden-Unruh/Data-Deliverable-Tool").toURI());
						} catch (Exception e1) {
							updateInfo(InfoText.DESKTOP);
						}
					else
						updateInfo(InfoText.DESKTOP);
				}
			}
		});

		run.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				if (checkCorrectSelections()) {
					SwingWorker<Boolean, Void> sw = new SwingWorker<Boolean, Void>() {
						@Override
						protected Boolean doInBackground() throws Exception {

							init();
							renameSheets();
							loadSheets();
							buildingValidation();
							towerValidation();
							groundsValidation();
							siteInventory();
							terminate();

							run.setEnabled(true);
							return true;
						}

						@Override
						protected void done() {
							try {
								get();
							} catch (InterruptedException e) {
								e.printStackTrace();
							} catch (ExecutionException e) {
								e.getCause().printStackTrace();
								String[] choices = { "Close", "More Info..." };
								updateInfo(InfoText.ERROR);
								run.setEnabled(true);
								if (JOptionPane.showOptionDialog(window,
										String.format("Unexpected Problem:\n%s", e.getCause().toString()), "Error",
										JOptionPane.DEFAULT_OPTION, JOptionPane.ERROR_MESSAGE, null, choices,
										choices[0]) == 1) {
									StringWriter sw = new StringWriter();
									e.printStackTrace(new PrintWriter(sw));
									JTextArea jta = new JTextArea(25, 50);
									jta.setText(String.format("Full Error Stack Trace:\n%s", sw.toString()));
									jta.setEditable(false);
									JOptionPane.showMessageDialog(window, new JScrollPane(jta), "Error",
											JOptionPane.ERROR_MESSAGE);
								}
							}
						}
					};
					run.setEnabled(false);
					sw.execute();
				} else
					updateInfo(InfoText.SELECT_PROMPT);
			}
		});

		window.pack();
		window.setVisible(true);
	}
	
	/**
	 * Loads the files into program-accessible format
	 * @throws FileNotFoundException if Files are not found - shouldn't happen, they're from file selection windows
	 * @throws IOException if there's an error reading the files
	 */
	private static void loadSheets() throws FileNotFoundException, IOException {
		updateInfo(InfoText.LOAD_SHEETS);
		deliverableBook = new XSSFWorkbook(new FileInputStream(selectedFiles[0]));
		workbookBook = new XSSFWorkbook(new FileInputStream(selectedFiles[1]));
	}

	/**
	 * Runs building validation section of program - see github readme for details
	 * @throws IOException if there's an error writing to the info file - only writes if location number isn't found in workbook
	 */
	private static void buildingValidation() throws IOException {
		XSSFSheet buildingSheet = deliverableBook.getSheet("Building Validation");

		// make a highlighted red style
		XSSFCellStyle redHighlight = buildingSheet.getRow(0).getCell(0).getCellStyle();
		redHighlight.setFillBackgroundColor(new HSSFColor(0, 0, Color.RED));

		int rows = buildingSheet.getPhysicalNumberOfRows();
		for (int i = 1; i < rows; i++) {
			XSSFRow activeRow = buildingSheet.getRow(i);
			String location = activeRow.getCell(3).toString();
			XSSFRow workbookRow = getCorrespondingRow(workbookBook.getSheet("BTG Validation"), location, 2);
			if (workbookRow != null) {
				setCell(workbookRow, 14, activeRow, 12);
				setCell(workbookRow, 3, activeRow, 13);
				setCell(workbookRow, 6, activeRow, 18);
				setCell(workbookRow, 7, activeRow, 19);
				setCell(workbookRow, 8, activeRow, 20);
				setCell(workbookRow, 9, activeRow, 23);
				setCell(workbookRow, 10, activeRow, 24);
				setCell(workbookRow, 11, activeRow, 32);
				setCell(workbookRow, 13, activeRow, 37);
				try {
					int deliverableSF = Integer.parseInt(activeRow.getCell(28).toString());
					int workbookSF = Integer.parseInt(workbookRow.getCell(4).toString());
					if (deliverableSF != workbookSF) {
						activeRow.getCell(28).setCellStyle(redHighlight);
					}
				} catch (NumberFormatException e) {
				}
			} else
				writeToInfo.append(String.format("Location number not found in workbook: %s\n", location));
		}
	}
	
	/**
	 * Runs tower validation section of program - see github readme for details
	 * @throws IOException if there's an error writing to info file - only writes if location number not found in workbook
	 */
	private static void towerValidation() throws IOException {
		XSSFSheet towerSheet = deliverableBook.getSheet("Tower Validation");

		int rows = towerSheet.getPhysicalNumberOfRows();
		for (int i = 1; i < rows; i++) {
			XSSFRow activeRow = towerSheet.getRow(i);
			String location = activeRow.getCell(3).toString();
			XSSFRow workbookRow = getCorrespondingRow(workbookBook.getSheet("BTG Validation"), location, 2);
			if (workbookRow != null) {
				setCell(workbookRow, 14, activeRow, 9);
				setCell(workbookRow, 11, activeRow, 17);
			} else
				writeToInfo.append(String.format("Location number not found in workbook: %s\n", location));
		}
	}

	/**
	 * Runs tower validation section of program - see github readme for details
	 * @throws IOException if there's an error writing to info file - only writes if location number not found in workbook
	 */
	private static void groundsValidation() throws IOException {
		XSSFSheet groundsSheet = deliverableBook.getSheet("Grounds Validation");

		int rows = groundsSheet.getPhysicalNumberOfRows();
		for (int i = 1; i < rows; i++) {
			XSSFRow activeRow = groundsSheet.getRow(i);
			String location = activeRow.getCell(3).toString();
			XSSFRow workbookRow = getCorrespondingRow(workbookBook.getSheet("BTG Validation"), location, 2);
			if (workbookRow != null) {
				setCell(workbookRow, 14, activeRow, 9);
				setCell(workbookRow, 11, activeRow, 14);
			} else
				writeToInfo.append(String.format("Location number not found in workbook: %s\n", location));
		}
	}
	
	//TODO tank validation?
	/**
	 * Runs site inventory section of the program - see github readme for details
	 */
	private static void siteInventory() {
		XSSFSheet inventorySheet = deliverableBook.getSheet("Asset Validation");
		XSSFSheet workbookSheet = workbookBook.getSheet("Site Inventory");

		HashSet<Integer> rowsToCheck = new HashSet<>();
		for (int i = 1; i < workbookSheet.getPhysicalNumberOfRows(); i++)
			rowsToCheck.add(i);

		int rows = inventorySheet.getPhysicalNumberOfRows();
		for (int i = 1; i < rows; i++) {
			XSSFRow activeRow = inventorySheet.getRow(i);
			String maximoId = activeRow.getCell(1).toString();
			int workbookRowNum = getCorrespondingRowNumber(workbookSheet, maximoId, 50);

			if (workbookRowNum == -1) { // Maximo ID on deliverable, not in workbook
				activeRow.getCell(5).setCellValue("DECOMMISSIONED");
				continue;
			}

			XSSFRow workbookRow = workbookSheet.getRow(workbookRowNum); // Maximo IDs match on deliverable/workbook
			rowsToCheck.remove(workbookRowNum);
			if (workbookRow.getCell(8).toString().toLowerCase().equals("removed"))
				activeRow.getCell(5).setCellValue("DECOMMISSIONED");
			setCell(workbookRow, 15, activeRow, 13);
			setCell(workbookRow, 13, activeRow, 14);
			setCell(workbookRow, 27, activeRow, 15); // TODO confirm expected life = estimated service life
			activeRow.getCell(16).setCellValue(Integer.toString(Integer.parseInt(workbookRow.getCell(27).toString())
					+ Integer.parseInt(workbookRow.getCell(30).toString())));	//TODO confirm this is how I should do this
			setCell(workbookRow, 34, activeRow, 17);
		}
		
		Iterator<Integer> it = rowsToCheck.iterator();	// Item only on worksheet, not on delivarable TODO: duplicate rows if Maximo ID blank
		int counter = 1;
		while (it.hasNext()) {
			int currentRow = inventorySheet.getPhysicalNumberOfRows();
			Integer i = it.next();
			XSSFRow newRow = inventorySheet.getRow(currentRow);
			XSSFRow prevRow = inventorySheet.getRow(currentRow - 1);
			XSSFRow workbookRow = workbookSheet.getRow(i);
			
			String assetName = workbookRow.getCell(7).toString();
			String buildingName = workbookRow.getCell(3).toString();
			
			setCell(prevRow, 0, newRow, 0);
			newRow.getCell(1).setCellValue(Integer.toString(counter++) + "NEW");
			newRow.getCell(2).setCellValue(JOptionPane.showInputDialog(window, String.format("What is the location ID for the asset titled %s in the building titled %s?", assetName, buildingName), JOptionPane.QUESTION_MESSAGE));
			newRow.getCell(4).setCellValue(assetName);
			newRow.getCell(5).setCellValue("OPERATING");
			newRow.getCell(6).setCellValue(""); //TODO: where does usage come from?
			newRow.getCell(7).setCellValue("FACILITIES");
			setCell(prevRow, 8, newRow, 8);
			setCell(workbookRow, 46, newRow, 9);
			newRow.getCell(12).setCellValue(newRow.getCell(0).toString().substring(3)); //TODO confirm this is an appropriate way to get inspection date
			setCell(workbookRow, 13, newRow, 15);
		}
	}
	
	/**
	 * Sets the specified cell on a given row to the contents of the specified cell on another given row
	 * @param readRow the row to read from
	 * @param readCol the index of cell of readRow to read from
	 * @param writeRow the row to write to
	 * @param writeCol the index of cell on writeRow to write to
	 */
	private static void setCell(XSSFRow readRow, int readCol, XSSFRow writeRow, int writeCol) {
		writeRow.getCell(writeCol).setCellValue(readRow.getCell(readCol).toString()); // TODO something if cell isn't
																						// empty/doesn't match new value
	}
	
	/**
	 * Gets the row from a given sheet that contains the specified String in its cell with specified index
	 * @param sheet the sheet to read from
	 * @param value the String to find
	 * @param matchCol the index of the cell within the returned Row
	 * @return the row that contains the given string in the specified cell
	 * @see #getCorrespondingRowNumber(XSSFSheet, String, int)
	 */
	private static XSSFRow getCorrespondingRow(XSSFSheet sheet, String value, int matchCol) {
		int num = getCorrespondingRowNumber(sheet, value, matchCol);
		if (num > -1)
			return sheet.getRow(num);
		return null;
	}

	/**
	 * Gets the index of the row from a given sheet that contains the specified String in its cell with specified index
	 * @param sheet the sheet to read from
	 * @param value the String to find
	 * @param matchCol the indesx of the cell within the returned Row
	 * @return the index of the row that contains the given string in the specified cell
	 * @see #getCorrespondingRow(XSSFSheet, String, int)
	 */
	private static int getCorrespondingRowNumber(XSSFSheet sheet, String value, int matchCol) {
		int rows = sheet.getPhysicalNumberOfRows();
		for (int i = 0; i < rows; i++)
			if (sheet.getRow(i).getCell(matchCol).toString().equals(value))
				return i;
		return -1;
	}

	/**
	 * Mapping of old to new sheet names
	 */
	private final static HashMap<String, String> nameMap = new HashMap<String, String>();
	/**
	 * Renames sheets according to the specified mapping - found in file <a href="file:/../resources/newNames.dat">newNames.dat</a>
	 * @throws IOException if there's an error opening the .dat file
	 */
	private static void renameSheets() throws IOException {
		for (int i = 0; i < deliverableBook.getNumberOfSheets(); i++) {
			String sheetName = deliverableBook.getSheetName(i);
			if (nameMap.containsKey(sheetName))
				deliverableBook.setSheetName(i, nameMap.get(sheetName));
			else
				writeToInfo.append(String.format("Sheet name not found in map: %s\n", sheetName));
		}
	}
	
	/**
	 * Checks if the user has selected two valid `.xlsx` files
	 * @return true if the user has selected valid files
	 */
	private static boolean checkCorrectSelections() {
		return FileNameUtils.getExtension(selectedFiles[0].getName()).equals("xlsx") && FileNameUtils.getExtension(selectedFiles[1].getName()).equals("xlsx");
	}
	
	/**
	 * Creates a GridBagConstraints object with the given attributes, and all other values set to defaults
	 * @param x horizontal location in grid bag
	 * @param y vertical location in grid bag
	 * @param width columns spanned in grid bag
	 * @param height rows spanned in grid bag
	 * @return the new GridBagConstraints object
	 */
	private static GridBagConstraints simpleConstraints(int x, int y, int width, int height) {
		return new GridBagConstraints(x, y, width, height, 0, 0, GridBagConstraints.CENTER, 0, new Insets(0, 0, 0, 0),
				0, 0);
	}
	
	/**
	 * Updates info to the specified enum value
	 * @param text the value to set to
	 */
	static void updateInfo(InfoText text) {
		infoText = text;
		info.setText(getInfoText());
		window.pack();
	}

	/**
	 * Gets the String corresponding to the current value of infoText
	 * @return the corresponding String to infoTexts value
	 */
	static String getInfoText() {
		switch (infoText) {
		case ERROR:
			return "Error encountered";
		case SELECT_PROMPT:
			return "Select 2 files of type .xlsx to continue";
		case DESKTOP:
			return "Help tried to open a browser, but failed. This could be due to a security restriction.";
		case LOAD_SHEETS:
			return "Opening spreadsheets...";
		case INIT:
			return "Initializing";
		}
		return null;
	}
}
