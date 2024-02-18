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

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

enum InfoText {
	ERROR, SELECT_PROMPT, DESKTOP, LOAD_SHEETS, INIT
}

public class Main {

	static JFrame window;

	static File[] selectedFiles = new File[2];

	static JLabel info = new JLabel();

	static XSSFWorkbook deliverableBook, workbookBook;

	static InfoText infoText;
	
	static FileWriter writeToInfo;

	public static void main(String[] args) {
		openWindow();
	}

	static final Pattern RENAME_LINE_PATTERN = Pattern.compile("(.+),(.+)");

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
		File infoFile = new File(String.format("%s\\Deliverable Program Information %s.txt", selectedFiles[0].getParent(), dateTime));
		infoFile.createNewFile();
		writeToInfo = new FileWriter(infoFile);
		writeToInfo.append(String.format("Info log for deliverable tool run %s\nBelow will be any potential errors encountered while running the tool. If it's empty, no errors were reported.\n", dateTime));
	}
	
	private static void terminate() throws IOException {
		writeToInfo.close();
	}

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

	private static void loadSheets() throws FileNotFoundException, IOException {
		updateInfo(InfoText.LOAD_SHEETS);
		deliverableBook = new XSSFWorkbook(new FileInputStream(selectedFiles[0]));
		workbookBook = new XSSFWorkbook(new FileInputStream(selectedFiles[1]));
	}

	private static void buildingValidation() throws IOException {
		// Add inspection date to column M in all rows
		XSSFSheet buildingSheet = deliverableBook.getSheet("Building Validation");
		
		// make a highlighted red style
		XSSFCellStyle redHighlight = buildingSheet.getRow(0).getCell(0).getCellStyle();
		redHighlight.setFillBackgroundColor(new HSSFColor(0, 0, Color.RED));
		
		int rows = buildingSheet.getPhysicalNumberOfRows();
		for (int i = 1; i < rows; i++) {
			XSSFRow activeRow = buildingSheet.getRow(i);
			String location = activeRow.getCell(3).toString();
			XSSFRow workbookRow = getCorrespondingRow(workbookBook.getSheet("BTG Validation"), location, 2);
			if(workbookRow != null) {
				setCell(workbookRow, 14, activeRow, 12);
				setCell(workbookRow, 3, activeRow, 13);
				setCell(workbookRow, 6, activeRow, 18);
				setCell(workbookRow, 7, activeRow, 19);
				setCell(workbookRow, 8, activeRow, 20);
				setCell(workbookRow, 9, activeRow, 23);
				setCell(workbookRow, 10, activeRow, 24);
				setCell(workbookRow, 11, activeRow, 32);
				setCell(workbookRow, 13, activeRow, 37);
				
				int deliverableSF = Integer.parseInt(activeRow.getCell(28).toString());
				int workbookSF = Integer.parseInt(workbookRow.getCell(4).toString());
				if (deliverableSF != workbookSF) {
					activeRow.getCell(28).setCellStyle(redHighlight);
				}
			} else
				writeToInfo.append(String.format("Location number not found in workbook: %s\n", location));
		}
	}

	private static void towerValidation() {

	}

	private static void groundsValidation() {

	}

	private static void siteInventory() {

	}
	
	private static void setCell(XSSFRow readRow, int readCol, XSSFRow writeRow, int writeCol) {
		writeRow.getCell(writeCol).setCellValue(readRow.getCell(readCol).toString()); //TODO something if cell isn't empty/doesn't match new value
	}
	
	private static XSSFRow getCorrespondingRow(XSSFSheet sheet, String value, int matchCol) {
		int rows = sheet.getPhysicalNumberOfRows();
		for (int i = 0; i < rows; i++)
			if(sheet.getRow(i).getCell(matchCol).toString().equals(value))
				return sheet.getRow(i);
		return null;
	}

	private final static HashMap<String, String> nameMap = new HashMap<String, String>();

	private static void renameSheets() throws IOException {
		for (int i = 0; i < deliverableBook.getNumberOfSheets(); i++) {
			String sheetName = deliverableBook.getSheetName(i);
			if (nameMap.containsKey(sheetName))
				deliverableBook.setSheetName(i, nameMap.get(sheetName));
			else
				writeToInfo.append(String.format("Sheet name not found in map: %s\n", sheetName));
		}
	}

	private static boolean checkCorrectSelections() {
		// TODO
		return true;
	}

	private static GridBagConstraints simpleConstraints(int x, int y, int width, int height) {
		return new GridBagConstraints(x, y, width, height, 0, 0, GridBagConstraints.CENTER, 0, new Insets(0, 0, 0, 0),
				0, 0);
	}

	static void updateInfo(InfoText text) {
		infoText = text;
		info.setText(getInfoText());
		window.pack();
	}

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
