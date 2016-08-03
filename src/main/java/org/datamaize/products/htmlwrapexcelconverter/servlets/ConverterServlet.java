/**
 * 
 */
package org.datamaize.products.htmlwrapexcelconverter.servlets;

/**
 * @author ashwinrayaprolu
 *
 */
import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.datamaize.products.htmlwrapexcelconverter.Converter;

@SuppressWarnings("serial")
@WebServlet(urlPatterns = { "/convert" })
public class ConverterServlet extends HttpServlet {
	static Logger logger = java.util.logging.Logger.getLogger(ConverterServlet.class.getName());

	/* (non-Javadoc)
	 * @see javax.servlet.http.HttpServlet#doGet(javax.servlet.http.HttpServletRequest, javax.servlet.http.HttpServletResponse)
	 */
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {

		try {
			File xlFile = new File(request.getParameter("filePath"));
			Date now = new Date();
			String currDate = new SimpleDateFormat("yyyy_MM_dd_HH_mm_ss").format(now);
			// fileName = "SampleExcel_" + currDate;
			String fileName = currDate + xlFile.getName();

			response.setContentType("application/ms-excel");
			response.setHeader("Expires:", "0"); // eliminates browsercaching
			response.setHeader("Content-Disposition", "attachment; filename=" + fileName + ".xls");
			int length = Converter.convert(request.getParameter("filePath"), response.getOutputStream());
			response.setContentLength(length);
		} catch (Exception e) {
			logger.log(Level.SEVERE, "Error Converting HTML to XLS ", e);
		} finally {
			try {
				response.getOutputStream().flush();
			} catch (IOException e) {
				e.printStackTrace();
			}

		}

	}
}
