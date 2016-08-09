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
import java.util.Enumeration;
import java.util.List;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.lang.StringUtils;
import org.datamaize.products.htmlwrapexcelconverter.ConversionContext;
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

			String fileName = "";
			Date now = new Date();
			
			
			String requestPath = request.getParameter("filePath");
			String replaceContent = request.getParameter("replaceContent");
			ConversionContext context = new ConversionContext();
			context.setContentToReplace(replaceContent);
			String currDate = new SimpleDateFormat("yyyy_MM_dd_HH_mm_ss").format(now);
			// fileName = "SampleExcel_" + currDate;
			int length = Converter.convert(requestPath, response.getOutputStream(),context,response);
			
			//response.setContentType("application/ms-excel");
			//response.setHeader("Content-Transfer-Encoding", "binary");
			//response.setHeader("Expires:", "0"); // eliminates browsercaching
			//response.setHeader("Content-Disposition", "attachment; filename=" + fileName + ".xls");
			response.setContentLength(length);
			if (StringUtils.startsWithIgnoreCase(requestPath, "http")) {
				fileName = context.getFileName();
			} else {
				File xlFile = new File(requestPath);
				fileName = currDate + xlFile.getName();
			}
			
			if(StringUtils.isBlank(fileName)){
				fileName = "data";
			}
			
			System.out.println("Rendering Response for:"+fileName);
			
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
