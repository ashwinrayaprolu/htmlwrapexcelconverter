package org.datamaize.products.htmlwrapexcelconverter;

public class ConversionContext {
	private String fileName = "";
	private String contentToReplace = "";

	/**
	 * @return the fileName
	 */
	public String getFileName() {
		return fileName;
	}

	/**
	 * @param fileName
	 *            the fileName to set
	 */
	public void setFileName(String fileName) {
		this.fileName = fileName;
	}

	/**
	 * @return the contentToReplace
	 */
	public String getContentToReplace() {
		return contentToReplace;
	}

	/**
	 * @param contentToReplace the contentToReplace to set
	 */
	public void setContentToReplace(String contentToReplace) {
		this.contentToReplace = contentToReplace;
	}
	
	

}
