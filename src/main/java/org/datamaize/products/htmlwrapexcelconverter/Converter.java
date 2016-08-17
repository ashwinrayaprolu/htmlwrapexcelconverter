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
import org.apache.commons.lang.WordUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;
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

			String htmlData = FileUtils.readFileToString(new File("/Users/ashwinrayaprolu/Desktop/sa_export_local.xls"));
			//System.out.println(stripInvalidTables(htmlData));
			
			//System.out.println("$.000.0".replaceAll("\\$", ""));
			
			
			ConversionContext context = new ConversionContext();
			convert("/Users/ashwinrayaprolu/Desktop/sa_export_local.xls", fileOutputStream, context, null);

			// convert("http://172.16.40.60/WebPortal/reports/cal/tmp/AuditCheck5087.asp",fileOutputStream,context);
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

	public static int convertURL(String urlPath, OutputStream out, ConversionContext context, HttpServletResponse response) {
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

		if (response != null) {
			if (StringUtils.isBlank(fileName)) {
				fileName = "data.xls";
			}
			response.setContentType("application/ms-excel");
			response.setHeader("Content-Transfer-Encoding", "binary");
			response.setHeader("Expires:", "0"); // eliminates browsercaching
			response.setHeader("Content-Disposition", "attachment; filename=" + fileName);
		}

		htmlData = stripInvalidTables(htmlData);
		// System.out.println(htmlData);
		HSSFWorkbook wb = null;
		try {
			// create work book
			wb = new HSSFWorkbook();
			// create excel sheet for page 1
			HSSFSheet sheet = wb.createSheet();

			// Set Header Font
			HSSFFont headerFont = wb.createFont();
			headerFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
			headerFont.setFontHeightInPoints((short) 12);

			// Set Header Style
			CellStyle headerStyle = wb.createCellStyle();
			headerStyle.setFillBackgroundColor(IndexedColors.BLACK.getIndex());
			headerStyle.setAlignment(CellStyle.ALIGN_CENTER);
			headerStyle.setFont(headerFont);
			headerStyle.setBorderTop(HSSFCellStyle.BORDER_MEDIUM);
			headerStyle.setBorderBottom(HSSFCellStyle.BORDER_MEDIUM);

			CellStyle numberStyle = wb.createCellStyle();
			numberStyle.setFillBackgroundColor(IndexedColors.BLACK.getIndex());
			numberStyle.setAlignment(CellStyle.ALIGN_RIGHT);
			// numberStyle.setBorderTop(HSSFCellStyle.BORDER_MEDIUM);
			// numberStyle.setBorderBottom(HSSFCellStyle.BORDER_MEDIUM);
			numberStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("#,##0.00"));

			int rowCount = 0;

			/*
			Folllowing code parse html table
			*/
			Document doc = Jsoup.parse(htmlData);

			generateExcel(wb, sheet, headerStyle, numberStyle, rowCount, doc);

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
			try {
				wb.close();
			} catch (IOException e) {
				// e.printStackTrace();
			}
		}

		return 0;

	}

	private static String stripInvalidTables(String input) {
		StringBuffer output = new StringBuffer();
		input = input.toLowerCase();
		int startIndex = 0;
		int tableStartIndex = 0;
		int nextTableStartIndex = 0;
		int tableEndIndex = 0;

		while ((tableStartIndex = StringUtils.indexOf(input, "<table", startIndex)) != -1) {
			tableEndIndex = StringUtils.indexOf(input, "</table", tableStartIndex);
			nextTableStartIndex = StringUtils.indexOf(input, "<table", tableStartIndex + 1);

			if (nextTableStartIndex == -1) {
				output.append(input.substring(startIndex));
				break;
			} else if (nextTableStartIndex < tableEndIndex) {
				startIndex = nextTableStartIndex;
				// Wrong case to dont append this
			} else {
				output.append(input.substring(startIndex, nextTableStartIndex));
				startIndex = nextTableStartIndex;
			}

		}

		return output.toString();
	}

	/**
	 * @param filePath
	 */
	public static int convert(String filePath, OutputStream out, ConversionContext context, HttpServletResponse response) {
		HSSFWorkbook wb = null;
		try {
			if (StringUtils.startsWithIgnoreCase(filePath, "http")) {
				return convertURL(filePath, out, context, response);
			}

			String htmlData = FileUtils.readFileToString(new File(filePath));

			if (StringUtils.isNotBlank(context.getContentToReplace())) {
				htmlData = htmlData.replaceAll(context.getContentToReplace(), "");
			}

			// create work book
			wb = new HSSFWorkbook();
			// create excel sheet for page 1
			HSSFSheet sheet = wb.createSheet();

			// Set Header Font
			HSSFFont headerFont = wb.createFont();
			headerFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
			headerFont.setFontHeightInPoints((short) 12);

			// Set Header Style
			CellStyle headerStyle = wb.createCellStyle();
			headerStyle.setFillBackgroundColor(IndexedColors.BLACK.getIndex());
			headerStyle.setAlignment(CellStyle.ALIGN_CENTER);
			headerStyle.setFont(headerFont);
			headerStyle.setBorderTop(HSSFCellStyle.BORDER_MEDIUM);
			headerStyle.setBorderBottom(HSSFCellStyle.BORDER_MEDIUM);

			CellStyle numberStyle = wb.createCellStyle();
			numberStyle.setFillBackgroundColor(IndexedColors.BLACK.getIndex());
			numberStyle.setAlignment(CellStyle.ALIGN_RIGHT);
			// numberStyle.setBorderTop(HSSFCellStyle.BORDER_MEDIUM);
			// numberStyle.setBorderBottom(HSSFCellStyle.BORDER_MEDIUM);
			numberStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("#,##0.00"));

			int rowCount = 0;
			Row header;

			/*
			Folllowing code parse html table
			*/
			Document doc = Jsoup.parse(htmlData);

			generateExcel(wb, sheet, headerStyle, numberStyle, rowCount, doc);

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
			try {
				wb.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}

		return 0;

	}

	/**
	 * @param wb
	 * @param sheet
	 * @param headerStyle
	 * @param numberStyle
	 * @param rowCount
	 * @param doc
	 */
	private static void generateExcel(HSSFWorkbook wb, HSSFSheet sheet, CellStyle headerStyle, CellStyle numberStyle, int rowCount, Document doc) {
		Row header;
		int colCount = 0;
		Cell cell;
		for (Element table : doc.select("table")) {

			// loop through all tr of table
			for (Element row : table.select("tr")) {
				rowCount++;
				// create row for each
				header = sheet.createRow(rowCount);
				// loop through all tag of Row
				Elements ths = row.select("th");
				int count = 0;
				for (Element element : ths) {
					// set header style
					count++;
					cell = header.createCell(count);
					cell.setCellValue(element.text());
					cell.setCellStyle(headerStyle);

					// Merges the cells
					String colspan = element.attr("colspan");
					if (StringUtils.isNotBlank(colspan)) {
						try {
							boolean isMerged = false;
							List<CellRangeAddress> mergedList = sheet.getMergedRegions();
							for (CellRangeAddress cellRange : mergedList) {
								if (cellRange.getFirstRow() <= cell.getRowIndex() && cellRange.getLastRow() >= cell.getRowIndex()
										&& cellRange.getFirstColumn() <= cell.getColumnIndex() && cellRange.getLastColumn() >= cell.getColumnIndex()) {
									isMerged = true;
								}
							}

							if (!isMerged) {
								sheet.addMergedRegion(new CellRangeAddress(cell.getRowIndex(), cell.getRowIndex(), cell.getColumnIndex(), cell.getColumnIndex()
										+ Integer.parseInt(colspan) - 1));
							}
						} catch (Exception e) {
							e.printStackTrace();
						}
					}

				}
				// now loop through all td tag
				Elements tds = row.select("td");
				count = 0;
				colCount = tds.size();
				for (Element element : tds) {
					// create cell for each tag
					count++;
					cell = header.createCell(count);
					if (element.hasClass("headr")) {
						cell.setCellStyle(headerStyle);
					}
					if (StringUtils.startsWith(element.text().trim(), "0") && !StringUtils.contains(element.text(), ".") && !StringUtils.contains(element.text(), "$")) {
						cell.setCellValue(WordUtils.capitalizeFully(element.text()));
					} else {
						try {
							double doubleVal = Double.parseDouble(element.text().replaceAll(",", "").replaceAll("\\$", ""));
							cell.setCellValue(doubleVal);
							cell.setCellStyle(numberStyle);
						} catch (Exception e) {
							cell.setCellValue(WordUtils.capitalizeFully(element.text()));
						}
					}

					// Merges the cells
					String colspan = element.attr("colspan");
					if (StringUtils.isNotBlank(colspan)) {
						try {

							boolean isMerged = false;
							List<CellRangeAddress> mergedList = sheet.getMergedRegions();
							for (CellRangeAddress cellRange : mergedList) {
								if (cellRange.getFirstRow() <= cell.getRowIndex() && cellRange.getLastRow() >= cell.getRowIndex()
										&& cellRange.getFirstColumn() <= cell.getColumnIndex() && cellRange.getLastColumn() >= cell.getColumnIndex()) {
									isMerged = true;
								}
							}

							if (!isMerged) {
								sheet.addMergedRegion(new CellRangeAddress(cell.getRowIndex(), cell.getRowIndex(), cell.getColumnIndex(), cell.getColumnIndex()
										+ Integer.parseInt(colspan) - 1));
							}
						} catch (Exception e) {
							e.printStackTrace();
						}
					}

				}

			}

			// set auto size column for excel sheet
			sheet = wb.getSheetAt(0);
			for (int j = 0; j < colCount; j++) {
				sheet.autoSizeColumn(j);
			}

		}
	}
}
