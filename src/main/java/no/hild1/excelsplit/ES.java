package no.hild1.excelsplit;

import java.awt.Container;
import java.awt.Dimension;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.util.Date;
import java.util.Map;

import javax.swing.Box;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;


public class ES {
	JFileChooser fc;
	Container panel;
	JLabel inputFileLabel;
	JButton runButton;
	boolean file = false;
	boolean nummer = true;
	JFrame guiFrame;
	TeePrintStream outStream;
	private Workbook inputWorkbook = null;
	private DataFormatter formatter = null;
	private FormulaEvaluator evaluator = null;
	
	protected boolean splitExcel(File inputFile) throws FileNotFoundException, InvalidFormatException, IOException {
		if (!inputFile.exists()) {
			throw new IllegalArgumentException("Klarte ikke finne " + inputFile);
		}
		this.openWorkbook(inputFile);
		this.splitWorkbook();
		System.out.println("Finished splitting");
		return true;
	}
	
	
	private boolean splitWorkbook() throws InvalidFormatException, IOException {
		Sheet sheet = null;

		System.out.println("Starter splitting av fil");
		int numSheets = this.inputWorkbook.getNumberOfSheets();
		boolean foundAtleastOneValid = false;

		for(int i = 0; i < numSheets; i++) {
			sheet = this.inputWorkbook.getSheetAt(i);
			String name = sheet.getSheetName();
			if (name.contains("Ark1")) {
				System.out.println("Fant Ark1, prosesserer");
				foundAtleastOneValid = true;
				processSheet(sheet);
				System.out.println("Finished prosessing.");
			} else {
				System.out.println("Fant ukjent regneark: " + name + ", ignorerer");
			}
		}
		if (!foundAtleastOneValid) {
			throw new InvalidFormatException("Fant ikke innland-regneark, ingen output laget. (utland-konvertering støttes ikke atm.)");
		}
		return true;
	}


	private void processSheet(Sheet sheet) throws IOException {
		// Er det 2 eller flere rows (dvs minst header + 1) i arket? 
		if(sheet.getPhysicalNumberOfRows() >= 2) { 
			int lastRowNum = 0;
			Row header = null;
			
			lastRowNum = sheet.getLastRowNum(); // hent siste row det er skrevet i
			
			System.out.println("Regnearket har " + lastRowNum + " rader" );
			
			header = sheet.getRow(0);
			
			int lastCellNum = header.getLastCellNum();
			
			System.out.println("Header har " + lastCellNum + " kolonner" );
			
			String header1 = text(header, 0);
			String header2 = text(header, 1);
			String header3 = text(header, 2);
			String header4 = text(header, 3);
			
			if (header1.equals("Header 1") && header2.equals("Header 2") && header3.equals("Header 3") && header4.equals("Header 4")) {
				System.out.println("Første rad ser OK ut, fortsetter");
				Row row = null;
				
			    Map<String, XSSFWorkbook> header2types = null;
				for(int j = 1; j <= lastRowNum; j++) {
					System.out.println("Prosesserer rad "  + j + " av " + lastRowNum);
					row = sheet.getRow(j);
					handleRow(row, j, header, header2types);
				}
				for(Map.Entry<String, XSSFWorkbook> entry: header2types.entrySet()) {
					FileOutputStream out = new FileOutputStream("Some_name_" + entry.getKey() + ".xlss");
					entry.getValue().write(out);
					out.close();
				}
			}
		}
	}
	
	private void handleRow(Row row, int j, Row header, Map<String, XSSFWorkbook> header2types) throws IOException {
		int HEADER1 = 0, HEADER2 = 1, HEADER3 = 2, HEADER4 = 3;
		String header2forthisrow = text(row, HEADER2);
		XSSFWorkbook w = null;
		Sheet s = null;
		Row r = null;
	    if (!header2types.containsKey(header2forthisrow)) {
	    	w = new XSSFWorkbook();
	    	s = w.createSheet();
	    	r = s.createRow(0);
	    	// insert "header" into "r" somehow
	    	header2types.put(header2forthisrow, w);
		} else {
			w = header2types.get(header2forthisrow);
			s = w.getSheetAt(0);
		}
	    r = s.createRow(s.getLastRowNum() + 1);  
    	// insert data "row" into "r" somehow
	}


	private String text(Row row, int pos) {
		Cell cell = row.getCell(pos);
		if (cell != null) {
			String s = null;
			switch (cell.getCellType()) {
			case Cell.CELL_TYPE_NUMERIC:
				s = String.format("%1.0f", cell.getNumericCellValue());
				break;

			case Cell.CELL_TYPE_FORMULA:
				s = this.formatter.formatCellValue(cell, this.evaluator);
				break;
			default:
				s = cell.getStringCellValue();
			}
			return s.trim();
		} else {
			return "";
		}
	}

	private void openWorkbook(File file) throws FileNotFoundException, IOException, InvalidFormatException {
		FileInputStream fis = null;
		try {
			System.out.println("Åpner arbeidsbok [" + file.getName() + "]");
			fis = new FileInputStream(file);
			this.inputWorkbook = WorkbookFactory.create(fis);
			this.evaluator = this.inputWorkbook.getCreationHelper().createFormulaEvaluator();
			this.formatter = new DataFormatter(true);
		}
		finally {
			if(fis != null) {
				fis.close();
			}
		}
	}


	public ES() {
		guiFrame = new JFrame();
		panel = guiFrame.getContentPane();
		// make sure the program exits when the frame closes
		guiFrame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
		guiFrame.setTitle("This Is A Name");
		guiFrame.setSize(450, 250);

		// This will center the JFrame in the middle of the screen
		guiFrame.setLocationRelativeTo(null);
		panel.setLayout(new GridBagLayout());
		GridBagConstraints c = new GridBagConstraints();

		fc = new JFileChooser(System.getProperty("user.dir"));
		fc.addChoosableFileFilter(new XSLXFilter());
		fc.setFileFilter(new XSLXFilter());
		fc.setAcceptAllFileFilterUsed(false);

		runButton = new JButton("Konverter");
		c.insets = new Insets(3, 3, 3, 3);
		c.fill = GridBagConstraints.HORIZONTAL;
		c.gridwidth = 2;
		c.gridx = 0;
		c.gridy = 0;

		c.gridx = 2; // one right

		JButton velgFilButton = new JButton("Velg fil");
		velgFilButton.setMinimumSize(new Dimension(250, 10));
		inputFileLabel = new JLabel("Ingen fil valgt");

		c.gridx = 0; // back left
		c.gridwidth = 4;
		c.gridy++; // one down
		c.gridy++; // one down
		panel.add(inputFileLabel, c);
		c.gridy++; // one down
		panel.add(new JLabel(), c);
		c.gridy++; // one down
		panel.add(new JLabel(), c);
		c.gridy++; // one down
		panel.add(velgFilButton, c);
		c.gridy++; // one down
		panel.add(new JLabel(), c);
		c.gridy++; // one down
		panel.add(new JLabel(), c);
		c.gridy++; // one down
		panel.add(new JLabel(), c);
		c.gridy++; // one down
		runButton.setEnabled(false);
		panel.add(runButton, c);
		c.gridwidth = 4;
		c.gridy++; // one down
		panel.add(runButton, c);

		JMenuBar menuBar = new JMenuBar();

		JMenu hjelpMenu = new JMenu("Hjelp");
		JMenuItem hjelpMenuItem = new JMenuItem("Hjelp");
		hjelpMenuItem.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent arg0) {
				JOptionPane.showMessageDialog(null, Data.HELPTEXT, "Hjelp",
						JOptionPane.INFORMATION_MESSAGE);
			}
		});

		JMenuItem omMenuItem = new JMenuItem("Om");
		omMenuItem.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent e) {
				JOptionPane.showMessageDialog(null, Data.LICENSE, "Om",
						JOptionPane.INFORMATION_MESSAGE);
			}
		});
		hjelpMenu.add(hjelpMenuItem);
		hjelpMenu.add(omMenuItem);
		menuBar.add(Box.createHorizontalGlue());
		menuBar.add(hjelpMenu);
		guiFrame.setJMenuBar(menuBar);

		runButton.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent event) {
				File inputFile = new File((fc.getSelectedFile())
						.getAbsolutePath());

				// System.out.println(TimeZone.getDefault().get);
				String logfileName = inputFile.getParent() + File.separator
						+ String.format("%tFT%<tR", new Date()) + ".log";
				File logFile = new File(logfileName);
				PrintStream printStream;
				try {
					printStream = new PrintStream(new FileOutputStream(logFile));
					outStream = new TeePrintStream(System.out, printStream);
					System.out.println("Logging to stdout and " + logfileName);
					System.setOut(outStream);
					if (splitExcel(inputFile)) {
						System.out.println("Foo!!");
						JOptionPane
								.showMessageDialog(
										null,
										Data.FINISHED
												+ "\n\nDette vinduet vil lukke seg når du klikker OK",
										"Ferdig",
										JOptionPane.INFORMATION_MESSAGE);
						guiFrame.dispose();
					}

				} catch (Exception e) {
					System.out.print(e);
					e.printStackTrace(System.out);
					JOptionPane.showMessageDialog(null, "Noe feil skjedde. Se "
							+ logFile.getAbsolutePath() + " for detaljer",
							"Ops!", JOptionPane.ERROR_MESSAGE);
					guiFrame.dispose();
				}

			}
		});

		velgFilButton.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent event) {
				int returnVal = fc.showDialog(panel, "Konverter denne filen");
				if (returnVal == JFileChooser.APPROVE_OPTION) {
					File inputFile = fc.getSelectedFile();
					inputFileLabel.setText(inputFile.getAbsolutePath());
					file = true;
					runButton.setEnabled((nummer && file));
					// repack to resize
					guiFrame.pack();
				} else {
					file = false;
					runButton.setEnabled((nummer && file));
				}

			}
		});

		// make sure the JFrame is visible
		guiFrame.setMinimumSize(new Dimension(300, 200));
		// guiFrame.pack();
		guiFrame.setVisible(true);
	}

	

	public static void main(String[] args) throws Exception {

		try {
			// Set System L&F
			UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
		} catch (UnsupportedLookAndFeelException e) {
			// handle exception
		} catch (ClassNotFoundException e) {
			// handle exception
		} catch (InstantiationException e) {
			// handle exception
		} catch (IllegalAccessException e) {
			// handle exception
		}

		try {
			new ES();
		} catch (Exception e) {
			System.out.println(e.getMessage());
			throw new Exception(e);
		}
	}
}
