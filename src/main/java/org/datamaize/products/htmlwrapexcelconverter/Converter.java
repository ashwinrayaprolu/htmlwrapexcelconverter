/**
 * 
 */
package org.datamaize.products.htmlwrapexcelconverter;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

/**
 * @author ashwinrayaprolu
 *
 */
public class Converter {
	static Logger logger = java.util.logging.Logger.getLogger(Converter.class.getName());

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		FileOutputStream fileOutputStream = null;
		try {
			fileOutputStream = new FileOutputStream("/Users/ashwinrayaprolu/test.xls");
			convert("/Users/ashwinrayaprolu/Desktop/AuditCheck.xls", fileOutputStream);

		} catch (Exception e) {
			logger.log(Level.SEVERE, "Error Converting HTML to XML ", e);
		} finally {
			try {
				fileOutputStream.flush();
			} catch (IOException e) {
				e.printStackTrace();
			}
			try {
				fileOutputStream.close();
			} catch (IOException e) {
				e.printStackTrace();
			}

		}
	}

	/**
	 * @param filePath
	 */
	public static int convert(String filePath, OutputStream out) {
		try {
			String htmlData = FileUtils.readFileToString(new File(filePath));
			// create work book
			HSSFWorkbook wb = new HSSFWorkbook();
			// create excel sheet for page 1
			HSSFSheet sheet = wb.createSheet();

			// Set Header Font
			HSSFFont headerFont = wb.createFont();
			headerFont.setBoldweight(headerFont.BOLDWEIGHT_BOLD);
			headerFont.setFontHeightInPoints((short) 12);

			// Set Header Style
			CellStyle headerStyle = wb.createCellStyle();
			headerStyle.setFillBackgroundColor(IndexedColors.BLACK.getIndex());
			headerStyle.setAlignment(headerStyle.ALIGN_CENTER);
			headerStyle.setFont(headerFont);
			headerStyle.setBorderTop(HSSFCellStyle.BORDER_MEDIUM);
			headerStyle.setBorderBottom(HSSFCellStyle.BORDER_MEDIUM);
			int rowCount = 0;
			Row header;

			/*
			Folllowing code parse html table
			*/
			Document doc = Jsoup.parse(htmlData);

			Cell cell;
			for (Element table : doc.select("table")) {
				rowCount++;
				// loop through all tr of table
				for (Element row : table.select("tr")) {
					// create row for each
					header = sheet.createRow(rowCount);
					// loop through all tag of Row
					Elements ths = row.select("th");
					int count = 0;
					for (Element element : ths) {
						// set header style
						cell = header.createCell(count);
						cell.setCellValue(element.text());
						cell.setCellStyle(headerStyle);
						count++;
					}
					// now loop through all td tag
					Elements tds = row.select("td.headr");
					count = 0;
					for (Element element : tds) {
						// create cell for each tag
						cell = header.createCell(count);
						cell.setCellStyle(headerStyle);
						cell.setCellValue(element.text());

						count++;
					}

					tds = row.select("td:not(.headr)");
					count = 0;
					for (Element element : tds) {
						// create cell for each tag
						cell = header.createCell(count);
						try {
							double doubleVal = Double.parseDouble(element.text().replaceAll(",", ""));
							cell.setCellValue(doubleVal);
						} catch (Exception e) {
							cell.setCellValue(element.text());
						}

						count++;
					}
					rowCount++;

					// set auto size column for excel sheet
					sheet = wb.getSheetAt(0);
					for (int j = 0; j < row.select("th").size(); j++) {
						sheet.autoSizeColumn(j);
					}
				}

			}

			ByteArrayOutputStream outByteStream = new ByteArrayOutputStream();
			wb.write(outByteStream);
			byte[] outArray = outByteStream.toByteArray();
			if (out != null) {
				out.write(outArray);
			}
			
			return outArray.length;

		} catch (Exception e) {
			e.printStackTrace();

		} finally {

		}
		
		return 0;

	}
}
