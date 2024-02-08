package com.myFileConverter.Converters.Utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.Reader;
import java.io.StringWriter;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.font.Standard14Fonts;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Element;
import com.itextpdf.text.Font;
import com.itextpdf.text.Font.FontFamily;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import com.opencsv.CSVReader;
import com.opencsv.exceptions.CsvException;
import com.opencsv.exceptions.CsvValidationException;

public class CsvToPDFConverter {

	public boolean convertToPdf(String inFilePath, String outFileDir) {

		boolean result = false;
		String inputCsvFile = inFilePath;
		String outPutFile = outFileDir + "\\" + getFileNameWithoutExtension(inFilePath) + ".pdf";

		try {
//        	try {
//				convertCSVToPlainTextPDFWithoutTables(inputCsvFile, outPutFile);
//			} catch (CsvValidationException e) {
//				// TODO Auto-generated catch block
//				e.printStackTrace();
//			} catch (DocumentException e) {
//				// TODO Auto-generated catch block
//				e.printStackTrace();
//			}
			File csvFile = new File(inputCsvFile);
			InputStream targetStream = new FileInputStream(csvFile);
			CSVReader reader = new CSVReader(new InputStreamReader(targetStream, "UTF-8"));
			Integer columns = 0;
			List<String[]> allLines = reader.readAll();
			if (!CollectionUtils.isEmpty(allLines)) {
				columns = Integer.valueOf(allLines.get(0).length);
				String fileNameWithoutExtension = getFileNameWithoutExtension(inFilePath);
				convertCSVToPdf(columns, allLines, fileNameWithoutExtension, outFileDir);
			}
			reader.close();
			result = true;
			System.out.println("PDF generated successfully: " + outPutFile);
		} catch (CsvValidationException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (Exception e) {
			System.out.println(e.getMessage());
			e.printStackTrace();
		}

		return result;
	}

	private static void convertCSVToPdf(Integer columns, List<String[]> reader, String fileName, String path) {

		Path location = Paths.get(path + File.separator + fileName);

		Document myPdf = new Document(PageSize.A4.rotate());
		try {
			PdfWriter.getInstance(myPdf, new FileOutputStream(location + ".pdf"));
			myPdf.open();
			PdfPTable table = new PdfPTable(columns);
			table.setWidthPercentage(100);
			table.getDefaultCell().setUseAscender(true);
			table.getDefaultCell().setUseDescender(true);
			if (!CollectionUtils.isEmpty(reader)) {
				reader.forEach(read -> {
					for (int i = 0; i < read.length; i++) {
						float fontSize = columns < 6 ? 12f : columns < 15 ? 8f : 3.5f;
						PdfPCell cells = new PdfPCell(new Phrase(read[i], new Font(FontFamily.HELVETICA, fontSize)));
//						cells.setColspan(5);
						cells.setHorizontalAlignment(Element.ALIGN_CENTER);
						cells.setVerticalAlignment(Element.ALIGN_LEFT);
						cells.setPadding(5.0f);
						table.addCell(cells);
					}
				});
				myPdf.add(table);
				myPdf.close();
			}
		} catch (DocumentException | IOException e) {
			System.out.println("Pdf writer creates error" + e);
		}
	}

	private static void convertCSVToPlainTextPDFWithoutTables(String csvFilePath, String pdfFilePath)
			throws IOException, DocumentException, CsvValidationException {
		Document document = new Document();
		PdfWriter.getInstance(document, new FileOutputStream(pdfFilePath));
		document.open();

		CSVReader reader = new CSVReader(new FileReader(csvFilePath));
		String[] header = reader.readNext();
		Paragraph paragraph = new Paragraph(String.join(", ", header));
		document.add(paragraph);

		String[] row;
		while ((row = reader.readNext()) != null) {
			paragraph = new Paragraph(String.join(", ", row));
			document.add(paragraph);
		}

		reader.close();
		document.close();
	}

	/**
	 * Returns filename only
	 * 
	 * @param filePath
	 * @return
	 */
	private static String getFileNameWithoutExtension(String filePath) {
		String fileName = "";
		try {
			File file = new File(filePath);
			if (file != null && file.exists()) {
				String name = file.getName();
				fileName = name.replaceFirst("[.][^.]+$", "");
			}
		} catch (Exception e) {
			e.printStackTrace();
			fileName = "";
		}

		return fileName;

	}

}
