package dataDeliverableTool;

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
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.net.URL;
import java.text.DecimalFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
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
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Possible values for infoText element within the Window of Data Deliverable
 * {@link Main#window}
 * 
 * @author Jaden Unruh
 * @since 0.0.1
 */
enum InfoText {
	ERROR, SELECT_PROMPT, DESKTOP, LOAD_SHEETS, INIT, CLOSING, BUILD_VALID, TOWER_VALID, GROUNDS_VALID, SITE_INV, DONE,
	DEF_DATA, COST_DATA, TANK_VALID, HEADERS
}

/**
 * Sheets of the deliverable file
 * 
 * @author Jaden
 * @since 1.1.0
 */
enum Sheet {
	BUILDING, TANK, TOWER, GROUNDS, ASSET, ORDERS, NEWORDERS, COSTDATA
}

/**
 * Primary class for Data Deliverable Tool.
 * 
 * @see <a href="https://github.com/Jaden-Unruh/data-deliverable-tool/">Data
 *      Deliverable Github</a>
 * @author Jaden Unruh
 * @since 0.0.1
 */
public class Main {

	/**
	 * Converts individual XSSFCell objects into the String excel shows to users
	 */
	static final DataFormatter FORMATTER = new DataFormatter();

	/**
	 * Primary GUI window
	 */
	static JFrame window;

	/**
	 * Files currently selected - updates every time the user presses 'ok' within a
	 * file selection prompt
	 */
	static File[] selectedFiles = new File[2];

	/**
	 * Info bar within {@link #window}
	 * 
	 * @see #infoText
	 */
	static JLabel info = new JLabel();

	/**
	 * Regular expression for location number: A##-## where A can be any capital
	 * letter, # can be any digit
	 */
	static final String LOCATION_REGEX = "^[A-Z]\\d{2}-\\d{2}$"; //$NON-NLS-1$

	/**
	 * Input field for location ID
	 */
	static EntryField locationEntry;

	/**
	 * Large objects representing the selected files - only update when the user
	 * selects 'run'
	 */
	static XSSFWorkbook deliverableBook, workbookBook;

	/**
	 * Current state of {@link #info}
	 * 
	 * @see #getInfoText()
	 */
	static InfoText infoText;

	/**
	 * Sheet we're currently completing the headers of
	 */
	static Sheet headerSheet = Sheet.BUILDING;

	/**
	 * Writer for generated info text file - only defined in {@link #init()}
	 */
	static FileWriter writeToInfo;

	/**
	 * Time the user clicks 'Run' in nanoseconds measured to some arbitrary moment -
	 * used to compute total run time
	 */
	static long startTime;

	/**
	 * Entry method for {@link Main}
	 * 
	 * @param args unused
	 */
	public static void main(String[] args) {
		openWindow();
	}

	/**
	 * Regex to pull old and new names from a line of
	 * <a href="file:/../resources/newNames.dat">newNames.dat</a>
	 */
	static final Pattern RENAME_LINE_PATTERN = Pattern.compile("(.+),(.+)"); //$NON-NLS-1$

	/**
	 * Regex to pull sheet name and column headers from a line of
	 * <a href="file:/../resources/columnHeaders.dat">columnHeaders.dat</a>
	 */
	static final Pattern COLUMN_HEADER_LINE_PATTERN = Pattern.compile("(.+):(.+)");

	/**
	 * Mapping of old to new sheet names in the deliverable file
	 */
	final static HashMap<String, String> nameMap = new HashMap<String, String>();

	/**
	 * Mapping of new sheet names to column header array
	 */
	final static HashMap<String, String[]> columnMap = new HashMap<String, String[]>();

	/**
	 * Initializes Data Deliverable tool - creates info file, pulls sheet renaming
	 * info
	 * 
	 * @throws IOException if creation of info file (and/or writer) fails, or
	 *                     grabbing sheet renaming .dat file fails
	 */
	private static void init() throws IOException {
		startTime = System.nanoTime();
		updateInfo(InfoText.INIT);
		try (BufferedReader reader = new BufferedReader(
				new InputStreamReader(Main.class.getResourceAsStream("newNames.dat")))) { //$NON-NLS-1$
			String[] lines = reader.lines().toArray(String[]::new);
			for (String line : lines) {
				Matcher match = RENAME_LINE_PATTERN.matcher(line);
				if (match.find())
					nameMap.put(match.group(1), match.group(2));
			}
		}

		try (BufferedReader reader = new BufferedReader(
				new InputStreamReader(Main.class.getResourceAsStream("columnHeaders.dat")))) { //$NON-NLS-1$
			String[] lines = reader.lines().toArray(String[]::new);
			for (String line : lines) {
				Matcher match = COLUMN_HEADER_LINE_PATTERN.matcher(line);
				if (match.find())
					columnMap.put(match.group(1), match.group(2).split(","));
			}
		}

		String dateTime = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd_HH.mm.ss.SSS")); //$NON-NLS-1$
		File infoFile = new File(
				String.format(Messages.getString("Main.infoFile.name"), selectedFiles[0].getParent(), dateTime)); //$NON-NLS-1$
		infoFile.createNewFile();
		writeToInfo = new FileWriter(infoFile);
		writeToInfo.append(String.format(Messages.getString("Main.infoFile.header"), //$NON-NLS-1$
				dateTime));
	}

	/**
	 * Safely saves and closes info file writer and spreadsheet files
	 * 
	 * @throws IOException if saving/closing fails
	 */
	private static void terminate() throws IOException {
		updateInfo(InfoText.CLOSING);
		FileOutputStream out = new FileOutputStream(new File(String.format("%s/%s (output).xlsx",
				selectedFiles[0].getParent(), FileNameUtils.getBaseName(selectedFiles[0].getName()))));
		deliverableBook.write(out);
		out.close();
		deliverableBook.close();
		workbookBook.close();

		String dateTime = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd_HH:mm:ss.SSS"));
		DecimalFormat secondFormatter = new DecimalFormat("#,###.########");

		double seconds = (double) (System.nanoTime() - startTime) / 1e9;

		writeToInfo.append(
				String.format(Messages.getString("Main.infoFile.footer"), dateTime, secondFormatter.format(seconds)));

		writeToInfo.close();
	}

	/**
	 * Opens and sets up the main window for the data deliverable tool
	 */
	private static void openWindow() {
		window = new JFrame(Messages.getString("Main.window.title")); //$NON-NLS-1$
		window.setLayout(new GridBagLayout());
		window.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);

		window.add(new JLabel(Messages.getString("Main.window.CAFilePrompt")), simpleConstraints(0, 0, 2, 1)); //$NON-NLS-1$
		JButton selectDeliverable = new SelectButton(0);
		window.add(selectDeliverable, simpleConstraints(2, 0, 1, 1));

		window.add(new JLabel(Messages.getString("Main.window.workbookFilePrompt")), simpleConstraints(0, 1, 2, 1)); //$NON-NLS-1$
		JButton selectWorkbook = new SelectButton(1);
		window.add(selectWorkbook, simpleConstraints(2, 1, 1, 1));

		window.add(new JLabel(Messages.getString("Main.window.locIDPrompt")), simpleConstraints(0, 2, 2, 1));

		locationEntry = new EntryField(LOCATION_REGEX, Messages.getString("Main.window.locIDDefText")); //$NON-NLS-1$
		window.add(locationEntry, simpleConstraints(2, 2, 1, 1));

		window.add(info, simpleConstraints(0, 3, 3, 1));

		JButton close = new JButton(Messages.getString("Main.window.close")); //$NON-NLS-1$
		window.add(close, simpleConstraints(0, 4, 1, 1));

		JButton help = new JButton(Messages.getString("Main.window.help")); //$NON-NLS-1$
		window.add(help, simpleConstraints(1, 4, 1, 1));

		final JButton run = new JButton(Messages.getString("Main.window.run")); //$NON-NLS-1$
		window.add(run, simpleConstraints(2, 4, 1, 1));

		close.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				System.exit(0);
			}
		});

		help.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				String[] choices = { Messages.getString("Main.window.help.close"), //$NON-NLS-1$
						Messages.getString("Main.window.help.github") }; //$NON-NLS-1$
				if (JOptionPane.showOptionDialog(window, Messages.getString("Main.window.help.text"), //$NON-NLS-1$
						Messages.getString("Main.window.help.title"), JOptionPane.DEFAULT_OPTION, //$NON-NLS-1$
						JOptionPane.INFORMATION_MESSAGE, null, choices, choices[0]) == 1) {
					if (Desktop.isDesktopSupported())
						try {
							Desktop.getDesktop()
									.browse(new URL("https://github.com/Jaden-Unruh/Data-Deliverable-Tool").toURI()); //$NON-NLS-1$
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

						protected Boolean doInBackground() throws Exception {

							init();
							loadSheets();
							renameSheets();
							completeHeaders();
							buildingValidation();
							towerValidation();
							groundsValidation();
							tankValidation();
							siteInventory();
							deficiencyData();
							costData();
							terminate();

							updateInfo(InfoText.DONE);

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
								String[] choices = { Messages.getString("Main.window.error.close"), //$NON-NLS-1$
										Messages.getString("Main.window.error.more") }; //$NON-NLS-1$
								updateInfo(InfoText.ERROR);
								run.setEnabled(true);
								if (JOptionPane.showOptionDialog(window,
										String.format(Messages.getString("Main.window.error.header"), //$NON-NLS-1$
												e.getCause().toString()),
										"Error", JOptionPane.DEFAULT_OPTION, JOptionPane.ERROR_MESSAGE, null, choices, //$NON-NLS-1$
										choices[0]) == 1) {
									StringWriter sw = new StringWriter();
									e.printStackTrace(new PrintWriter(sw));
									JTextArea jta = new JTextArea(25, 50);
									jta.setText(
											String.format(Messages.getString("Main.window.error.fst"), sw.toString())); //$NON-NLS-1$
									jta.setEditable(false);
									JOptionPane.showMessageDialog(window, new JScrollPane(jta), "Error", //$NON-NLS-1$
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
	 * 
	 * @throws FileNotFoundException if Files are not found - shouldn't happen,
	 *                               they're from file selection windows
	 * @throws IOException           if there's an error reading the files
	 */
	private static void loadSheets() throws FileNotFoundException, IOException {
		updateInfo(InfoText.LOAD_SHEETS);
		deliverableBook = new XSSFWorkbook(new FileInputStream(selectedFiles[0]));
		workbookBook = new XSSFWorkbook(new FileInputStream(selectedFiles[1]));
		deliverableBook.setMissingCellPolicy(Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
		workbookBook.setMissingCellPolicy(Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
	}

	/**
	 * Replaces all the headers for every sheet in the deliverable, correcting them
	 * to what's in columnHeaders.dat <br>
	 * Any data already in the deliverable will be kept, as long as it is under a
	 * correct header - even if it's in the wrong column number(letter). That is,
	 * data will be moved to the correct column as long as its header is one that is
	 * meant to be kept. Any data under a header that is not listed in the
	 * respective row of columnHeaders.dat will be lost in the output (it will
	 * remain in the input deliverable, which will be left untouched)
	 * 
	 * @throws IOException if an attempted write to the info text file fails
	 */
	private static void completeHeaders() throws IOException {
		updateInfo(InfoText.HEADERS);

		updateHeaderSheet(Sheet.BUILDING);
		String buildingName = getSheetName(Sheet.BUILDING);
		completeHeadersOnSheet(deliverableBook.getSheet(buildingName), Arrays.asList(columnMap.get(buildingName)));

		updateHeaderSheet(Sheet.TANK);
		String tankName = getSheetName(Sheet.TANK);
		completeHeadersOnSheet(deliverableBook.getSheet(tankName), Arrays.asList(columnMap.get(tankName)));

		updateHeaderSheet(Sheet.TOWER);
		String towerName = getSheetName(Sheet.TOWER);
		completeHeadersOnSheet(deliverableBook.getSheet(towerName), Arrays.asList(columnMap.get(towerName)));

		updateHeaderSheet(Sheet.GROUNDS);
		String groundsName = getSheetName(Sheet.GROUNDS);
		completeHeadersOnSheet(deliverableBook.getSheet(groundsName), Arrays.asList(columnMap.get(groundsName)));

		updateHeaderSheet(Sheet.ASSET);
		String assetName = getSheetName(Sheet.ASSET);
		completeHeadersOnSheet(deliverableBook.getSheet(assetName), Arrays.asList(columnMap.get(assetName)));

		updateHeaderSheet(Sheet.ORDERS);
		String ordersName = getSheetName(Sheet.ORDERS);
		completeHeadersOnSheet(deliverableBook.getSheet(ordersName), Arrays.asList(columnMap.get(ordersName)));

		updateHeaderSheet(Sheet.NEWORDERS);
		String newOrdersName = getSheetName(Sheet.NEWORDERS);
		completeHeadersOnSheet(deliverableBook.getSheet(newOrdersName), Arrays.asList(columnMap.get(newOrdersName)));

		updateHeaderSheet(Sheet.COSTDATA);
		String costDataName = getSheetName(Sheet.COSTDATA);
		completeHeadersOnSheet(deliverableBook.getSheet(costDataName), Arrays.asList(columnMap.get(costDataName)));
	}

	/**
	 * Replaces the headers on a specific sheet to match the provided list of
	 * headers, maintaining and moving data to stay in the proper columns.
	 * 
	 * @param sheet           the sheet to replace the headers of
	 * @param sheetNewColumns the new headers to use
	 * @throws IOException if an attempted write to the info text file fails
	 * @see Main#completeHeaders()
	 */
	private static void completeHeadersOnSheet(XSSFSheet sheet, List<String> sheetNewColumns) throws IOException {
		HashMap<String, LinkedList<String>> contentsByColumn = new HashMap<String, LinkedList<String>>();

		int currentNumCols = getNumberOfColumns(sheet.getRow(0));
		int currentNumRows = getNumberOfRows(sheet, 0);

		for (int col = 0; col < currentNumCols; col++) {
			String firstCell = FORMATTER.formatCellValue(sheet.getRow(0).getCell(col));
			if (sheetNewColumns.contains(firstCell)) {
				LinkedList<String> thisColumn = new LinkedList<>();
				for (int row = 1; row < currentNumRows; row++)
					thisColumn.offer(FORMATTER.formatCellValue(sheet.getRow(row).getCell(col)));
				contentsByColumn.put(firstCell, thisColumn);
			} else
				writeToInfo.append(String.format(Messages.getString("Main.infoFile.extraCol"), firstCell,
						getSheetName(headerSheet)));
		}

		// remove all contents
		for (int i = 0; i < currentNumRows; i++)
			sheet.createRow(i);

		for (int col = 0; col < sheetNewColumns.size(); col++) {
			String header = sheetNewColumns.get(col);
			sheet.getRow(0).getCell(col).setCellValue(header);
			LinkedList<String> contents;
			if ((contents = contentsByColumn.get(header)) != null) {
				int currentRow = 1;
				while (contents.peek() != null)
					sheet.getRow(currentRow++).getCell(col).setCellValue(contents.poll());
			}
		}
	}

	/**
	 * Gets the number of columns in the row, counting only non-empty cells and
	 * stopping at the first empty cell
	 * 
	 * @param row the row to count the columns in
	 * @return the number of columns
	 * @see Main#getNumberOfRows(XSSFSheet, int)
	 */
	private static int getNumberOfColumns(XSSFRow row) {
		int ret = -1;
		while (FORMATTER.formatCellValue(row.getCell(++ret)).length() != 0)
			;
		return ret;
	}

	/**
	 * Gets the number of rows in the sheet, counting only non-empty cells in the
	 * given column and stopping at the first empty cell
	 * 
	 * @param sheet the sheet to count the rows of
	 * @param col   the (0-)index of the column to count in
	 * @return the number of rows
	 * @see Main#getNumberOfColumns(XSSFRow)
	 */
	private static int getNumberOfRows(XSSFSheet sheet, int col) {
		int defRows = sheet.getPhysicalNumberOfRows();
		for (int i = 0; i < defRows; i++)
			if (FORMATTER.formatCellValue(sheet.getRow(i).getCell(col)).length() == 0)
				return i + 1;
		return defRows;
	}

	/**
	 * Runs building validation section of program - see github readme for details
	 * 
	 * @throws IOException if there's an error writing to the info file - only
	 *                     writes if location number isn't found in workbook
	 */
	private static void buildingValidation() throws IOException {
		updateInfo(InfoText.BUILD_VALID);
		XSSFSheet buildingSheet = deliverableBook.getSheet(getSheetName(Sheet.BUILDING));

		// make a highlighted red style
		XSSFCellStyle redHighlight = buildingSheet.getRow(0).getCell(0).getCellStyle();
		final byte[] RED = { Byte.MAX_VALUE, 0, 0 };
		redHighlight.setFillBackgroundColor(new XSSFColor(RED));

		int rows = buildingSheet.getPhysicalNumberOfRows();
		for (int i = 1; i < rows; i++) {
			XSSFRow activeRow = buildingSheet.getRow(i);
			String location = activeRow.getCell(3).toString();
			XSSFRow workbookRow = getCorrespondingRow(
					workbookBook.getSheet(Messages.getString("Main.sheetName.workbook.btgValidation")), location, //$NON-NLS-1$
					2);
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
				writeToInfo.append(String.format(Messages.getString("Main.infoFile.locNumNotFound"), location)); //$NON-NLS-1$
		}
	}

	/**
	 * Runs tower validation section of program - see github readme for details
	 * 
	 * @throws IOException if there's an error writing to info file - only writes if
	 *                     location number not found in workbook
	 */
	private static void towerValidation() throws IOException {
		updateInfo(InfoText.TOWER_VALID);
		XSSFSheet towerSheet = deliverableBook.getSheet(getSheetName(Sheet.TOWER));

		int rows = towerSheet.getPhysicalNumberOfRows();
		for (int i = 1; i < rows; i++) {
			XSSFRow activeRow = towerSheet.getRow(i);
			String location = activeRow.getCell(3).toString();
			XSSFRow workbookRow = getCorrespondingRow(
					workbookBook.getSheet(Messages.getString("Main.sheetName.workbook.btgValidation")), location, 2); //$NON-NLS-1$
			if (workbookRow != null) {
				setCell(workbookRow, 14, activeRow, 9);
				setCell(workbookRow, 11, activeRow, 17);
			} else
				writeToInfo.append(String.format(Messages.getString("Main.infoFile.locNumNotFound"), location)); //$NON-NLS-1$
		}
	}

	/**
	 * Runs tower validation section of program - see github readme for details
	 * 
	 * @throws IOException if there's an error writing to info file - only writes if
	 *                     location number not found in workbook
	 */
	private static void groundsValidation() throws IOException {
		updateInfo(InfoText.GROUNDS_VALID);
		XSSFSheet groundsSheet = deliverableBook.getSheet(getSheetName(Sheet.GROUNDS));

		int rows = groundsSheet.getPhysicalNumberOfRows();
		for (int i = 1; i < rows; i++) {
			XSSFRow activeRow = groundsSheet.getRow(i);
			String location = activeRow.getCell(3).toString();
			XSSFRow workbookRow = getCorrespondingRow(
					workbookBook.getSheet(Messages.getString("Main.sheetName.workbook.btgValidation")), location, 2); //$NON-NLS-1$
			if (workbookRow != null) {
				setCell(workbookRow, 14, activeRow, 9);
				setCell(workbookRow, 11, activeRow, 14);
			} else
				writeToInfo.append(String.format(Messages.getString("Main.infoFile.locNumNotFound"), location)); //$NON-NLS-1$
		}
	}

	/**
	 * Runs tank validation section of program - see github readme for details
	 * 
	 * @throws IOException if there's an error writing to info file - only writes if
	 *                     location number not found in workbook
	 */
	private static void tankValidation() throws IOException {
		updateInfo(InfoText.TANK_VALID);
		XSSFSheet groundsSheet = deliverableBook.getSheet(getSheetName(Sheet.TANK));

		int rows = groundsSheet.getPhysicalNumberOfRows();
		for (int i = 1; i < rows; i++) {
			XSSFRow activeRow = groundsSheet.getRow(i);
			String location = activeRow.getCell(3).toString();
			XSSFRow workbookRow = getCorrespondingRow(
					workbookBook.getSheet(Messages.getString("Main.sheetName.workbook.btgValidation")), location, 2); //$NON-NLS-1$
			if (workbookRow != null) {
				setCell(workbookRow, 14, activeRow, 9);
				setCell(workbookRow, 11, activeRow, 14);
			} else
				writeToInfo.append(String.format(Messages.getString("Main.infoFile.locNumNotFound"), location)); //$NON-NLS-1$
		}
	}

	/**
	 * Runs site inventory section of the program - see github readme for details
	 */
	private static void siteInventory() {
		updateInfo(InfoText.SITE_INV);
		XSSFSheet inventorySheet = deliverableBook.getSheet(getSheetName(Sheet.ASSET));
		XSSFSheet workbookSheet = workbookBook.getSheet(Messages.getString("Main.sheetName.workbook.siteInventory")); //$NON-NLS-1$

		HashSet<Integer> rowsToCheck = new HashSet<>();
		for (int i = 1; i < workbookSheet.getPhysicalNumberOfRows(); i++)
			if (FORMATTER.formatCellValue(workbookSheet.getRow(i).getCell(0)).length() != 0)
				rowsToCheck.add(i);

		int rows = inventorySheet.getPhysicalNumberOfRows();
		for (int i = 1; i < rows; i++) {
			XSSFRow activeRow = inventorySheet.getRow(i);
			String maximoId = activeRow.getCell(1).toString();
			int workbookRowNum = getCorrespondingRowNumber(workbookSheet, maximoId, 50);

			if (workbookRowNum == -1) { // Maximo ID on deliverable, not in workbook
				activeRow.getCell(5).setCellValue(Messages.getString("Main.sheet.decommissionedText")); //$NON-NLS-1$
				continue;
			}

			XSSFRow workbookRow = workbookSheet.getRow(workbookRowNum); // Maximo IDs match on deliverable/workbook
			rowsToCheck.remove(workbookRowNum);
			if (workbookRow.getCell(8).toString().toLowerCase().equals(Messages.getString("Main.sheet.removedText"))) //$NON-NLS-1$
				activeRow.getCell(5).setCellValue(Messages.getString("Main.sheet.decommissionedText")); //$NON-NLS-1$
			setCell(workbookRow, 15, activeRow, 13);
			setCell(workbookRow, 13, activeRow, 14);
			setCell(workbookRow, 27, activeRow, 15);
			activeRow.getCell(16)
					.setCellValue(Integer.toString((int) (Double.parseDouble(workbookRow.getCell(27).toString())
							+ Double.parseDouble(workbookRow.getCell(30).toString()))));
			setCell(workbookRow, 34, activeRow, 17);
		}

		Iterator<Integer> it = rowsToCheck.iterator(); // Item only on worksheet, not on delivarable
		String locID = locationEntry.getText();
		int counter = 1;
		while (it.hasNext()) {
			int currentRow = inventorySheet.getPhysicalNumberOfRows();
			Integer i = it.next();
			XSSFRow newRow = inventorySheet.createRow(currentRow);
			XSSFRow prevRow = inventorySheet.getRow(currentRow - 1);
			XSSFRow workbookRow = workbookSheet.getRow(i);

			String assetName = FORMATTER.formatCellValue(workbookRow.getCell(7));
			String buildingName = FORMATTER.formatCellValue(workbookRow.getCell(3));

			setCell(prevRow, 0, newRow, 0);
			newRow.getCell(1)
					.setCellValue(Integer.toString(counter++) + Messages.getString("Main.sheet.newAssetIdSuffix")); //$NON-NLS-1$
			newRow.getCell(2).setCellValue(Messages.getString("Main.sheet.IAFMS_LOC_ID")); // TODO building location ID
																							// - not overall loc id,
																							// Hoku will add new
			// column
			newRow.getCell(4).setCellValue(assetName);
			newRow.getCell(5).setCellValue(Messages.getString("Main.sheet.operatingText")); //$NON-NLS-1$
			newRow.getCell(6).setCellValue(Messages.getString("Main.sheet.usage")); // TODO: where does //$NON-NLS-1$
																					// usage come
																					// from?
			newRow.getCell(7).setCellValue(Messages.getString("Main.sheet.facilitiesText")); //$NON-NLS-1$
			setCell(prevRow, 8, newRow, 8);
			setCell(workbookRow, 46, newRow, 9);
			newRow.getCell(11).setCellValue(FORMATTER.formatCellValue(newRow.getCell(0)).substring(3));
			setCell(workbookRow, 15, newRow, 13);
			setCell(workbookRow, 14, newRow, 14);
			setCell(workbookRow, 27, newRow, 15);
			String cell27 = FORMATTER.formatCellValue(workbookRow.getCell(27)); //TODO this section is really ugly but it's late and I just want it to work so fix later
			String cell30 = FORMATTER.formatCellValue(workbookRow.getCell(30));
			newRow.getCell(16).setCellValue(
					Integer.toString((int) ((cell27.isEmpty() ? 0 : Double.parseDouble(FORMATTER.formatCellValue(workbookRow.getCell(27))))
							+ (cell30.isEmpty() ? 0 : Double.parseDouble(FORMATTER.formatCellValue(workbookRow.getCell(30)))))));
			setCell(workbookRow, 34, newRow, 17);
		}
	}

	/**
	 * Runs deficiency data section of the program - see github readme for details
	 */
	private static void deficiencyData() {
		updateInfo(InfoText.DEF_DATA);

		// Copy from building validation b/c deficiency data might be empty
		String inspNum = FORMATTER
				.formatCellValue(deliverableBook.getSheet(getSheetName(Sheet.BUILDING)).getRow(1).getCell(0));
		String siteID = FORMATTER
				.formatCellValue(deliverableBook.getSheet(getSheetName(Sheet.BUILDING)).getRow(1).getCell(2));

		XSSFSheet defSheet = deliverableBook.getSheet(getSheetName(Sheet.NEWORDERS));
		XSSFSheet workbookSheet = workbookBook.getSheet(Messages.getString("Main.sheetName.workbook.workItems"));

		// check how many rows Deficiency Data has:
		int startRow = 0;
		XSSFRow checkRow;
		while ((checkRow = defSheet.getRow(++startRow)) != null
				&& FORMATTER.formatCellValue(checkRow.getCell(0)).length() != 0)
			;

		int rows = workbookSheet.getPhysicalNumberOfRows();
		for (int i = 0; i < rows; i++) {

			XSSFRow activeRow = defSheet.createRow(i + startRow);
			XSSFRow workbookRow = workbookSheet.getRow(i + 1);

			if (FORMATTER.formatCellValue(workbookRow.getCell(0)).length() == 0)
				break;

			activeRow.getCell(0).setCellValue(inspNum);
			setCell(workbookRow, 7, activeRow, 1);
			setCell(workbookRow, 28, activeRow, 2);
			setCell(workbookRow, 27, activeRow, 3);
			activeRow.getCell(4).setCellValue(siteID);
			activeRow.getCell(7).setCellValue(Messages.getString("Main.sheet.deficiency.status"));
			setCell(workbookRow, 8, activeRow, 8);
			setCell(workbookRow, 13, activeRow, 10, 0, 1);
			setCell(workbookRow, 13, activeRow, 11, 2, 3);
			activeRow.getCell(12).setCellValue(Messages.getString("Main.sheet.deficiency.iaFunc"));
			setCell(workbookRow, 12, activeRow, 17, 0, 2);

			activeRow.getCell(19)
					.setCellValue(String.format("%s; %s", FORMATTER.formatCellValue(workbookRow.getCell(22)),
							FORMATTER.formatCellValue(workbookRow.getCell(23))));
		}
	}

	/**
	 * Runs cost data section of the program - see github readme for details
	 */
	private static void costData() {
		updateInfo(InfoText.COST_DATA);

		XSSFSheet costSheet = deliverableBook.getSheet(getSheetName(Sheet.COSTDATA));
		XSSFSheet workbookSheet = workbookBook.getSheet(Messages.getString("Main.sheetName.workbook.workItems"));

		String inspNum = FORMATTER
				.formatCellValue(deliverableBook.getSheet(getSheetName(Sheet.BUILDING)).getRow(1).getCell(0));
		String siteID = FORMATTER
				.formatCellValue(deliverableBook.getSheet(getSheetName(Sheet.BUILDING)).getRow(1).getCell(2));

		HashSet<Integer> rowsToAdd = new HashSet<>();
		int rows = workbookSheet.getPhysicalNumberOfRows();
		for (int i = 1; i < rows; i++)
			if (FORMATTER.formatCellValue(workbookSheet.getRow(i).getCell(0)).length() != 0
					&& getCorrespondingRowNumber(costSheet,
							FORMATTER.formatCellValue(workbookSheet.getRow(i).getCell(7)), 5) < 0)
				rowsToAdd.add(i);

		Iterator<Integer> it = rowsToAdd.iterator();

		while (it.hasNext()) {
			int copyRow = it.next();
			int currentRow = costSheet.getPhysicalNumberOfRows();
			XSSFRow activeRow = costSheet.createRow(currentRow);
			XSSFRow workbookRow = workbookSheet.getRow(copyRow);

			activeRow.getCell(0).setCellValue(inspNum);
			setCell(workbookRow, 7, activeRow, 1);
			setCell(workbookRow, 28, activeRow, 2);
			activeRow.getCell(3).setCellValue(siteID);
			activeRow.getCell(4).setCellValue(Messages.getString("Main.sheet.cost.type"));
			setCell(workbookRow, 7, activeRow, 5);
			activeRow.getCell(12).setCellValue(Messages.getString("Main.sheet.cost.lineType"));
			setCell(workbookRow, 26, activeRow, 15);
			setCell(workbookRow, 8, activeRow, 17);
		}
	}

	/**
	 * Sets the specified cell on a given row to the contents of a specified cell on
	 * another given row
	 * 
	 * @param readRow  the row to read from
	 * @param readCol  the index of cell of readRow to read from
	 * @param writeRow the row to write to
	 * @param writeCol the index of cell on writeRow to write to
	 * @see Main#setCell(XSSFRow, int, XSSFRow, int, int, int)
	 */
	private static void setCell(XSSFRow readRow, int readCol, XSSFRow writeRow, int writeCol) {
		writeRow.getCell(writeCol).setCellValue(FORMATTER.formatCellValue(readRow.getCell(readCol))); // TODO something
																										// if cell isn't
																										// empty/doesn't
																										// match new
																										// value
	}

	/**
	 * Sets the specified cell on a given row to a substring of the contents of a
	 * specified cell on another given row
	 * 
	 * @param readRow    the row to read from
	 * @param readCol    the index of cell of readRow to read from
	 * @param writeRow   the row to write to
	 * @param writeCol   the index of cell on writeRow to write to
	 * @param startIndex the start index for the substring of the string we're
	 *                   writing
	 * @param endIndex   the end index for the substring of the string we're writing
	 * @see Main#setCell(XSSFRow, int, XSSFRow, int)
	 * @see String#substring(int, int)
	 */
	private static void setCell(XSSFRow readRow, int readCol, XSSFRow writeRow, int writeCol, int startIndex,
			int endIndex) {
		writeRow.getCell(writeCol)
				.setCellValue(FORMATTER.formatCellValue(readRow.getCell(readCol)).substring(startIndex, endIndex));
	}

	/**
	 * Gets the row from a given sheet that contains the specified String in its
	 * cell with specified index
	 * 
	 * @param sheet    the sheet to read from
	 * @param value    the String to find
	 * @param matchCol the index of the cell within the returned Row
	 * @return the row that contains the given string in the specified cell
	 * @see #getCorrespondingRowNumber(XSSFSheet, String, int)
	 */
	private static XSSFRow getCorrespondingRow(XSSFSheet sheet, String value, int matchCol) {
		int num;
		return (num = getCorrespondingRowNumber(sheet, value, matchCol)) > -1 ? sheet.getRow(num) : null;
	}

	/**
	 * Gets the index of the row from a given sheet that contains the specified
	 * String in its cell with specified index
	 * 
	 * @param sheet    the sheet to read from
	 * @param value    the String to find
	 * @param matchCol the indesx of the cell within the returned Row
	 * @return the index of the row that contains the given string in the specified
	 *         cell
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
	 * Renames sheets according to the specified mapping - found in file
	 * <a href="file:/../resources/newNames.dat">newNames.dat</a>
	 * 
	 * @throws IOException if there's an error opening the .dat file
	 */
	private static void renameSheets() throws IOException {
		for (int i = 0; i < deliverableBook.getNumberOfSheets(); i++) {
			String sheetName = deliverableBook.getSheetName(i);
			if (nameMap.containsKey(sheetName))
				deliverableBook.setSheetName(i, nameMap.get(sheetName));
			else
				writeToInfo.append(String.format(Messages.getString("Main.infoFile.sheetNotFound"), sheetName)); //$NON-NLS-1$
		}
	}

	/**
	 * Checks if the user has selected two valid `.xlsx` files
	 * 
	 * @return true if the user has selected valid files
	 */
	private static boolean checkCorrectSelections() {
		return FileNameUtils.getExtension(selectedFiles[0].getName()).equals("xlsx") //$NON-NLS-1$
				&& FileNameUtils.getExtension(selectedFiles[1].getName()).equals("xlsx"); //$NON-NLS-1$
	}

	/**
	 * Creates a GridBagConstraints object with the given attributes, and all other
	 * values set to defaults
	 * 
	 * @param x      horizontal location in grid bag
	 * @param y      vertical location in grid bag
	 * @param width  columns spanned in grid bag
	 * @param height rows spanned in grid bag
	 * @return the new GridBagConstraints object
	 */
	private static GridBagConstraints simpleConstraints(int x, int y, int width, int height) {
		return new GridBagConstraints(x, y, width, height, 0, 0, GridBagConstraints.CENTER, 0, new Insets(0, 0, 0, 0),
				0, 0);
	}

	/**
	 * Updates info to the specified enum value
	 * 
	 * @param text the value to set to
	 */
	static void updateInfo(InfoText text) {
		infoText = text;
		info.setText(getInfoText());
		window.pack();
	}

	/**
	 * Gets the String corresponding to the current value of infoText
	 * 
	 * @return the corresponding String to infoTexts value
	 */
	static String getInfoText() {
		switch (infoText) {
		case ERROR:
			return Messages.getString("Main.infoText.error"); //$NON-NLS-1$
		case SELECT_PROMPT:
			return Messages.getString("Main.infoText.selectPrompt"); //$NON-NLS-1$
		case DESKTOP:
			return Messages.getString("Main.infoText.browserFail"); //$NON-NLS-1$
		case LOAD_SHEETS:
			return Messages.getString("Main.infoText.loadSheets"); //$NON-NLS-1$
		case INIT:
			return Messages.getString("Main.infoText.init"); //$NON-NLS-1$
		case BUILD_VALID:
			return Messages.getString("Main.infoText.buildingValidation"); //$NON-NLS-1$
		case CLOSING:
			return Messages.getString("Main.infoText.closing"); //$NON-NLS-1$
		case GROUNDS_VALID:
			return Messages.getString("Main.infoText.groundsValidation"); //$NON-NLS-1$
		case SITE_INV:
			return Messages.getString("Main.infoText.siteInventory"); //$NON-NLS-1$
		case TOWER_VALID:
			return Messages.getString("Main.infoText.towerValidation"); //$NON-NLS-1$
		case DONE:
			return Messages.getString("Main.infoText.done"); //$NON-NLS-1$
		case DEF_DATA:
			return Messages.getString("Main.infoText.deficiencyData"); //$NON-NLS-1$
		case COST_DATA:
			return Messages.getString("Main.infoText.costData"); //$NON-NLS-1$
		case TANK_VALID:
			return Messages.getString("Main.infoText.tankValidation"); //$NON-NLS-1$
		case HEADERS:
			return String.format(Messages.getString("Main.infoText.headers"), getSheetName(headerSheet));
		}
		return null;
	}

	/**
	 * Updates the GUI to show the sheet we're currently updating the headers of<br>
	 * Note that on reasonably fast computers and reasonably small data sets, these
	 * sheet names will likely switch faster than the user can see them - this is
	 * mostly in case something gets hung up, to know which sheet the hang happened
	 * on
	 * 
	 * @param sheet the sheet we're currently updating the headers of
	 */
	static void updateHeaderSheet(Sheet sheet) {
		headerSheet = sheet;
		updateInfo(infoText);
	}

	/**
	 * Fetches the String name associated with the given sheet from the messages.properties file
	 * @param sheet the sheet we want the name of
	 * @return the String sheet name
	 */
	static String getSheetName(Sheet sheet) {
		switch (sheet) {
		case BUILDING:
			return Messages.getString("Main.sheetName.deliverable.buildingValidation"); //$NON-NLS-1$
		case TANK:
			return Messages.getString("Main.sheetName.deliverable.tankValidation"); //$NON-NLS-1$
		case TOWER:
			return Messages.getString("Main.sheetName.deliverable.towerValidation"); //$NON-NLS-1$
		case GROUNDS:
			return Messages.getString("Main.sheetName.deliverable.groundsValidation"); //$NON-NLS-1$
		case ASSET:
			return Messages.getString("Main.sheetName.deliverable.assetValidation"); //$NON-NLS-1$
		case ORDERS:
			return Messages.getString("Main.sheetName.deliverable.workOrders"); //$NON-NLS-1$
		case NEWORDERS:
			return Messages.getString("Main.sheetName.deliverable.defData"); //$NON-NLS-1$
		case COSTDATA:
			return Messages.getString("Main.sheetName.deliverable.costData"); //$NON-NLS-1$
		}
		return null;
	}
}