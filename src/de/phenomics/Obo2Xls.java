package de.phenomics;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.stream.Collectors;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Option;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import ontologizer.go.Ontology;
import ontologizer.go.Term;
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
			parameterError = "Please provide the obo file using " + ontologyOpt.getOpt() + " or " + ontologyOpt.getLongOpt();
		}
		if (!(new File(oboFilePath).exists())) {
			parameterError = "obo file does not exist!";
		}

		/*
		 * Maybe something was wrong with the parameter. Print help for the user and die
		 * here...
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

		Term selectedRootTerm = null;
		if (classId != null) {
			classId = classId.trim();
			if (!classId.equals("")) {
				Term t = ontology.getTermIncludingAlternatives(classId);
				if (t != null) {
					selectedRootTerm = t;
				} else {
					System.err.println(
							"Warning! You selected " + classId + " as class to be investigated, but this ID couldn't be found in the ontology.");
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
		int rowIndex = 0;
		String[] headersTermAdd = new String[] { "Class Label", "Class ID", "Alternative IDs", "Synonyms (separated by semicolon)", "Definition",
				"Subclass-of (label+id)" };

		rowIndex = createHeaderRow(createHelper, style, rowIndex, sheet0, headersTermAdd);

		ArrayList<Term> termsToReport = new ArrayList<>();
		if (selectedRootTerm != null) {
			termsToReport.addAll(ontologySlim.getDescendants(selectedRootTerm));
		} else {
			termsToReport.addAll(ontology.getAllTerms());
		}

		for (Term term : termsToReport) {

			Row row = sheet0.createRow((short) rowIndex++);

			int columnIndex = 0;
			row.createCell(columnIndex++).setCellValue(createHelper.createRichTextString(term.getName()));
			row.createCell(columnIndex++).setCellValue(createHelper.createRichTextString(term.getIDAsString()));

			String altIds = Arrays.stream(term.getAlternatives()).map(Object::toString).collect(Collectors.joining("; "));
			row.createCell(columnIndex++).setCellValue(createHelper.createRichTextString(altIds));

			String synonyms = String.join("; ", term.getSynonymsArrayList());

			row.createCell(columnIndex++).setCellValue(createHelper.createRichTextString(synonyms));
			row.createCell(columnIndex++).setCellValue(createHelper.createRichTextString(term.getDefinition()));

			String parents = ontologySlim.getParents(term).stream().map(Object::toString).collect(Collectors.joining("; "));
			row.createCell(columnIndex++).setCellValue(createHelper.createRichTextString(parents));

		}

		for (int i = 0; i < wb.getNumberOfSheets(); i++) {
			wb.getSheetAt(i).setDefaultColumnWidth(defaultColumnWidth);
		}

		wb.write(fileOut);
		fileOut.close();
	}

	private static int createHeaderRow(XSSFCreationHelper createHelper, XSSFCellStyle style, int rowIndex, Sheet sheet, String[] strings) {
		Row headerrow = sheet.createRow((short) rowIndex++);
		int colIndex = 0;
		for (String s : strings) {
			headerrow.createCell(colIndex++).setCellValue(createHelper.createRichTextString(s));
		}
		for (int i = 0; i < headerrow.getLastCellNum(); i++) {
			Cell cell = headerrow.getCell(i);
			cell.setCellStyle(style);
		}
		return rowIndex;
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