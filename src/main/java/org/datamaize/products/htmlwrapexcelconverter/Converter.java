/**
 * 
 */
package org.datamaize.products.htmlwrapexcelconverter;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLConnection;
import java.util.List;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;

import javax.servlet.http.HttpServletResponse;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang.StringUtils;
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
			
			String htmlData = FileUtils.readFileToString(new File("/Users/ashwinrayaprolu/Desktop/AuditCheck.xls"));
			System.out.println(stripInvalidTables(htmlData));
			//convert("/Users/ashwinrayaprolu/Desktop/AuditCheck.xls",fileOutputStream);
			
			
			ConversionContext context = new ConversionContext();
			//convert("http://172.16.40.60/WebPortal/reports/cal/tmp/AuditCheck5087.asp",fileOutputStream,context);
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

	public static int convertTest(InputStream inp, OutputStream out, String fileName) {
		try {
			URL url = new URL("http://172.16.40.60/WebPortal/reports/cal/tmp/AuditCheck5087.asp");
			URLConnection con = url.openConnection();
			InputStream in = con.getInputStream();
			// get all headers
			Map<String, List<String>> map = con.getHeaderFields();
			for (Map.Entry<String, List<String>> entry : map.entrySet()) {
				if (StringUtils.containsIgnoreCase(entry.getKey(), "content-disposition")) {
					System.out.println("Key : " + entry.getKey() + " ,Value : " + entry.getValue());

					System.out.println("FileName:" + entry.getValue().toString().replaceAll("]", "").split("=")[1]);
				}
			}
			String encoding = con.getContentEncoding();
			encoding = encoding == null ? "UTF-8" : encoding;
			String body = IOUtils.toString(in, encoding);
			System.out.println(body);
		} catch (MalformedURLException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

		return 0;
	}

	public static int convertURL(String urlPath, OutputStream out,ConversionContext context,HttpServletResponse response ) {
		String fileName = "";
		String htmlData = "";
		try {
			// URL url = new
			// URL("http://172.16.40.60/WebPortal/reports/cal/tmp/AuditCheck5087.asp");
			URL url = new URL(urlPath);
			URLConnection con = url.openConnection();
			InputStream in = con.getInputStream();
			// get all headers
			Map<String, List<String>> map = con.getHeaderFields();
			for (Map.Entry<String, List<String>> entry : map.entrySet()) {
				if (StringUtils.containsIgnoreCase(entry.getKey(), "content-disposition")) {
					System.out.println("Key : " + entry.getKey() + " ,Value : " + entry.getValue());
					fileName = entry.getValue().toString().replaceAll("]", "").split("=")[1];
					context.setFileName(fileName);
					break;
				}
			}
			String encoding = con.getContentEncoding();
			encoding = encoding == null ? "UTF-8" : encoding;
			htmlData = IOUtils.toString(in, encoding);
			// System.out.println(body);
		} catch (MalformedURLException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		if(response!=null){
			response.setContentType("application/ms-excel");
			response.setHeader("Content-Transfer-Encoding", "binary");
			response.setHeader("Expires:", "0"); // eliminates browsercaching
			response.setHeader("Content-Disposition", "attachment; filename=" + fileName + ".xls");
		}
		
		htmlData = stripInvalidTables(htmlData);
		//System.out.println(htmlData);

		try {
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

			int colCount = 0;
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
					Elements tds = row.select("td");
					count = 0;
					colCount = tds.size();
					for (Element element : tds) {
						// create cell for each tag
						cell = header.createCell(count);
						if(element.hasClass("headr")){
							cell.setCellStyle(headerStyle);
						}
						
						try {
							double doubleVal = Double.parseDouble(element.text().replaceAll(",", ""));
							cell.setCellValue(doubleVal);
						} catch (Exception e) {
							cell.setCellValue(element.text());
						}

						count++;
					}
					

					
					rowCount++;

					
				}
				
				// set auto size column for excel sheet
				sheet = wb.getSheetAt(0);
				for (int j = 0; j < colCount; j++) {
					sheet.autoSizeColumn(j);
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

	
	private static String stripInvalidTables(String input){
		StringBuffer output = new StringBuffer();
		input = input.toLowerCase();
		int startIndex = 0;
		int tableStartIndex = 0;
		int nextTableStartIndex = 0;
		int tableEndIndex = 0;
		
		while((tableStartIndex = StringUtils.indexOf(input, "<table",startIndex) ) != -1){
			tableEndIndex = StringUtils.indexOf(input, "</table",tableStartIndex);
			nextTableStartIndex = StringUtils.indexOf(input, "<table",tableStartIndex+1);
			
			if(nextTableStartIndex== -1){
				output.append(input.substring(startIndex));
				break;
			}else if(nextTableStartIndex < tableEndIndex){
				startIndex = nextTableStartIndex;
				// Wrong case to dont append this
			}else{
				output.append(input.substring(startIndex, nextTableStartIndex));
				startIndex =nextTableStartIndex;
			}
			
			 
			
		}
		
		return output.toString();
	}
	
	/**
	 * @param filePath
	 */
	public static int convert(String filePath, OutputStream out,ConversionContext context,HttpServletResponse response) {
		try {
			if (StringUtils.startsWithIgnoreCase(filePath, "http")) {
				return convertURL(filePath, out,context,response);
			}

			String htmlData = FileUtils.readFileToString(new File(filePath));
			
			if(StringUtils.isNotBlank(context.getContentToReplace())){
				htmlData = htmlData.replaceAll(context.getContentToReplace(), "");
			}
			
			
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
