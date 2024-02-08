package com.myFileConverter;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;

import com.myFileConverter.Converters.Utils.CsvToPDFConverter;
import com.myFileConverter.Converters.Utils.ExcelToPDFConverter;
import com.myFileConverter.Converters.Utils.ImageToPDFConverter;
import com.myFileConverter.Converters.Utils.PptToPDFConverter;
import com.myFileConverter.Converters.Utils.WordToPDFConverter;

public class MyFileConverter {
	
//	public static void main(String[] args) {
//		File file = new File("C:\\Users\\saura\\Desktop\\experiment.txt");
//		try {
//			System.out.println(Files.readString(file.toPath()));
//		} catch (IOException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}
//		
//	}
	
	public static void main(String[] args) throws IOException {
//        File inputFile = new File("D:\\Downloads\\Experiment Folder\\Saurabh Resume - 30-10-2022.docx");
        File inputFile = new File("D:\\Downloads\\Experiment Folder\\Saurabh Resume - 30-10-2022.docx");
        File outputFile = new File(inputFile.getParentFile().getAbsolutePath());
        System.out.println("input File Exists - "+ inputFile.exists());
        System.out.println("inputFile abs path - " + inputFile.getAbsolutePath());
        System.out.println("outputFile abs path - " + outputFile.getAbsolutePath());
        
        convertToPdf(inputFile, outputFile);

        System.out.println("Conversion complete. Output saved to: " + outputFile.getAbsolutePath());
    }

    public static void convertToPdf(File inputFile, File outputFile) throws IOException {
        try {
            if (inputFile.getName().endsWith(".pptx")) {
            	// Handle .pptx conversion
                PptToPDFConverter pptToPDFConverter = new PptToPDFConverter();
                pptToPDFConverter.convertToPdf(inputFile.getAbsolutePath(), outputFile.getAbsolutePath(), inputFile.getName().split("\\.")[1]);
            } else if (inputFile.getName().endsWith(".ppt")) {
                // Handle .ppt conversion
            	PptToPDFConverter pptToPDFConverter = new PptToPDFConverter();
                pptToPDFConverter.convertToPdf(inputFile.getAbsolutePath(), outputFile.getAbsolutePath(), inputFile.getName().split("\\.")[1]);
            } else if (inputFile.getName().endsWith(".docx")) {
                WordToPDFConverter wordToPDFConverter = new WordToPDFConverter();
                wordToPDFConverter.convertToPdf(inputFile.getAbsolutePath(), outputFile.getAbsolutePath());
            } else if (inputFile.getName().endsWith(".doc")) {
                // Handle .doc conversion
            	WordToPDFConverter wordToPDFConverter = new WordToPDFConverter();
                wordToPDFConverter.convertToPdf(inputFile.getAbsolutePath(), outputFile.getAbsolutePath());
            } else if (inputFile.getName().endsWith(".xlsx")) {
                // Handle .xlsx conversion
            	ExcelToPDFConverter excelToPDFConverter = new ExcelToPDFConverter();
            	excelToPDFConverter.convertToPdf(inputFile.getAbsolutePath(), outputFile.getAbsolutePath());
            } else if (inputFile.getName().endsWith(".xls")) {
                // Handle .xls conversion
            	ExcelToPDFConverter excelToPDFConverter = new ExcelToPDFConverter();
            	excelToPDFConverter.convertToPdf(inputFile.getAbsolutePath(), outputFile.getAbsolutePath());
            } else if (inputFile.getName().endsWith(".txt")) {
                // Handle .txt conversion
            } else if (inputFile.getName().endsWith(".csv")) {
                // Handle .txt conversion
            	CsvToPDFConverter csvToPDFConverter = new CsvToPDFConverter();
            	csvToPDFConverter.convertToPdf(inputFile.getAbsolutePath(), outputFile.getAbsolutePath());
            } else if (
        			inputFile.getName().endsWith(".jpeg") 
        			|| inputFile.getName().endsWith(".jpg") 
        			|| inputFile.getName().endsWith(".png") 
        			|| inputFile.getName().endsWith(".gif") 
        			|| inputFile.getName().endsWith(".tif")
//        				To be handled in future
//        			|| inputFile.getName().endsWith(".psd") 
//        			|| inputFile.getName().endsWith(".svg")
        		  ) {
                // Handle image conversion
            	ImageToPDFConverter imageToPDFConverter = new ImageToPDFConverter();
            	imageToPDFConverter.convertToPdf(inputFile.getAbsolutePath(), outputFile.getAbsolutePath());
            } else {
                System.err.println("Unsupported file format: " + inputFile.getName());
                return;
            }
        } catch (Exception e) {
        	e.printStackTrace();
        	System.out.println(e.getMessage());
        }
    }

}
