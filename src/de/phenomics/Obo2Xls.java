package de.phenomics;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.stream.Collectors;

import ontologizer.go.Ontology;
import ontologizer.go.Term;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Option;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import sonumina.math.graph.SlimDirectedGraphView;
import util.OntologyUtil;

/**
 * 
 * 
 * 
 * @author Sebastian Koehler
 *
 */
public class Obo2Xls {

	private static int rowIndexNextRow;

	public static void main(String[] args) throws IOException, ParseException {

		// configure command line parser
		final CommandLineParser commandLineParser = new DefaultParser();
		HelpFormatter helpFormatter = new HelpFormatter();
		final Options options = new Options();

		// the obo file
		Option ontologyOpt = new Option("o", "ontology", true, "Path to obo file.");
		options.addOption(ontologyOpt);

		// the folder where the output is written to...
		Option classOpt = new Option("c", "class", true, "The ontology class for which the excel file shall be created.");
		options.addOption(classOpt);

		/*
		 * parse the command line
		 */
		final CommandLine commandLine = commandLineParser.parse(options, args);

		// get the pathes to files
		String oboFilePath = getOption(ontologyOpt, commandLine);
		String classId = getOption(classOpt, commandLine);

		// check that required parameters are set
		String parameterError = null;
		if (oboFilePath == null) {
			parameterError = "Please provide the obo file using -" + ontologyOpt.getOpt() + " or --" + ontologyOpt.getLongOpt();
		} else if (!(new File(oboFilePath).exists())) {
			parameterError = "obo file does not exist!";
		}

		/*
		 * Maybe something was wrong with the parameter. Print help for the user
		 * and die here...
		 */
		if (parameterError != null) {
			String className = Obo2Xls.class.getSimpleName();

			helpFormatter.printHelp(className, options);
			throw new IllegalArgumentException(parameterError);
		}

		/*
		 * Let's get started
		 */
		File oboFilePathFile = new File(oboFilePath);
		Ontology ontology = OntologyUtil.parseOntology(oboFilePathFile.getAbsolutePath());

		Term selectedRootTerm = ontology.getRootTerm();
		if (classId != null) {
			classId = classId.trim();
			if (!classId.equals("")) {
				Term t = ontology.getTermIncludingAlternatives(classId);
				if (t != null) {
					selectedRootTerm = t;
				} else {
					System.err.println("Warning! You selected " + classId
							+ " as class to be investigated, but this ID couldn't be found in the ontology.");
				}
			}
		}

		String fileName = oboFilePathFile.getName();
		String outfile = oboFilePath + ".xlsx";
		System.out.println("create xls version at " + outfile);
		createXlsFromObo(ontology, selectedRootTerm, fileName, outfile);

	}

	private static void createXlsFromObo(Ontology ontology, Term selectedRootTerm, String fileNameParsedOboFrom, String outfile) throws IOException {

		SlimDirectedGraphView<Term> ontologySlim = ontology.getSlimGraphView();

		String date = ontology.getTermMap().getDataVersion().replaceAll("releases/", "");

		int defaultColumnWidth = 25;
		XSSFWorkbook wb = new XSSFWorkbook();
		FileOutputStream fileOut = new FileOutputStream(outfile);
		XSSFCreationHelper createHelper = wb.getCreationHelper();

		XSSFFont bold = wb.createFont();
		bold.setBold(true);
		XSSFCellStyle style = wb.createCellStyle();
		style.setFont(bold);

		Sheet sheet0 = wb.createSheet("Excel version of " + fileNameParsedOboFrom + " version: " + date);
		String[] headersTermAdd = new String[] { "Class Label", "Class ID", "Alternative IDs", "Synonyms (separated by semicolon)", "Definition",
				"Subclass-of (label+id)" };

		rowIndexNextRow = 1;

		createHeaderRow(createHelper, style, sheet0, headersTermAdd);

		recursiveWriteTermsAndTheirChildren(selectedRootTerm, wb, sheet0, createHelper, ontologySlim, false, false);

		for (int i = 0; i < wb.getNumberOfSheets(); i++) {
			wb.getSheetAt(i).setDefaultColumnWidth(defaultColumnWidth);
		}

		wb.write(fileOut);
		fileOut.close();
	}

	private static void recursiveWriteTermsAndTheirChildren(Term currentTerm, XSSFWorkbook wb, Sheet sheet0, XSSFCreationHelper createHelper,
			SlimDirectedGraphView<Term> ontologySlim, boolean style, boolean isLast) {

		if (currentTerm.isObsolete())
			return;

		createRowForTerm(currentTerm, sheet0, createHelper, ontologySlim, wb, style);

		ArrayList<Term> children = ontologySlim.getChildren(currentTerm);
		for (int i = 0; i < children.size(); i++) {
			Term child = children.get(i);
			boolean isLastChild = (i == children.size() - 1);
			recursiveWriteTermsAndTheirChildren(child, wb, sheet0, createHelper, ontologySlim, !style, isLastChild);
		}
		if (children.size() < 1 && isLast)
			++rowIndexNextRow;
	}

	private static void createRowForTerm(Term term, Sheet sheet0, XSSFCreationHelper createHelper, SlimDirectedGraphView<Term> ontologySlim,
			XSSFWorkbook wb, boolean style) {

		Row row = sheet0.createRow(rowIndexNextRow);
		rowIndexNextRow++;
		int columnIndex = 0;

		XSSFCellStyle style1 = null;
		if (style) {
			style1 = wb.createCellStyle();
			style1.setFillForegroundColor(new XSSFColor(new java.awt.Color(220, 220, 220)));
			style1.setFillPattern(CellStyle.SOLID_FOREGROUND);
		}
		addCellWithStyle(term.getName(), createHelper, row, columnIndex++, style1);

		addCellWithStyle(term.getIDAsString(), createHelper, row, columnIndex++, style1);

		String altIds = Arrays.stream(term.getAlternatives()).map(Object::toString).collect(Collectors.joining("; "));
		addCellWithStyle(altIds, createHelper, row, columnIndex++, style1);

		String synonyms = String.join("; ", term.getSynonymsArrayList());
		addCellWithStyle(synonyms, createHelper, row, columnIndex++, style1);

		addCellWithStyle(term.getDefinition(), createHelper, row, columnIndex++, style1);

		String parents = ontologySlim.getParents(term).stream().map(Object::toString).collect(Collectors.joining("; "));
		addCellWithStyle(parents, createHelper, row, columnIndex++, style1);

	}

	private static void addCellWithStyle(String content, XSSFCreationHelper createHelper, Row row, int columnIndex, XSSFCellStyle style1) {
		Cell c = row.createCell(columnIndex);
		c.setCellValue(createHelper.createRichTextString(content));

		if (style1 != null) {
			c.setCellStyle(style1);
		}
	}

	private static void createHeaderRow(XSSFCreationHelper createHelper, XSSFCellStyle style, Sheet sheet, String[] strings) {
		Row headerrow = sheet.createRow((short) rowIndexNextRow);
		rowIndexNextRow++;
		int colIndex = 0;
		for (String s : strings) {
			headerrow.createCell(colIndex++).setCellValue(createHelper.createRichTextString(s));
		}
		for (int i = 0; i < headerrow.getLastCellNum(); i++) {
			Cell cell = headerrow.getCell(i);
			cell.setCellStyle(style);
		}
	}

	public static String getOption(Option opt, final CommandLine commandLine) {

		if (commandLine.hasOption(opt.getOpt())) {
			return commandLine.getOptionValue(opt.getOpt());
		}
		if (commandLine.hasOption(opt.getLongOpt())) {
			return commandLine.getOptionValue(opt.getLongOpt());
		}
		return null;
	}

}
