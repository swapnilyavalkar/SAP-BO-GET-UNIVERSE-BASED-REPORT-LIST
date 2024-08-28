//Authored By swapnil.yavalkar //

import java.io.File;
import java.io.FileOutputStream;
import java.util.StringTokenizer;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.crystaldecisions.sdk.exception.SDKException;
import com.crystaldecisions.sdk.framework.CrystalEnterprise;
import com.crystaldecisions.sdk.framework.IEnterpriseSession;
import com.crystaldecisions.sdk.framework.ISessionMgr;
import com.crystaldecisions.sdk.occa.infostore.IInfoStore;

public class InitCmsExcel {


	IEnterpriseSession es;
	protected static IInfoStore iStore;
	
	XSSFWorkbook workbook=null;
	Cell cell;
	XSSFFont font;
	XSSFCellStyle my_style_1;
	XSSFCellStyle my_style_2;
	XSSFSheet sheet;
	XSSFSheet sheetnotes;
	XSSFCellStyle my_style_0;
	int rownum;
	int rownumn;

	private String CMS;
	private String userId;
	private String password;
	private String filepath;	
	
	String universeid;
	String universename;
	private String finalfilename;
	
	protected static InitCmsExcel initcmsexcel = new InitCmsExcel();
	
	public void initCMSConnection() throws SDKException {
		System.out.println("======================================================================");
		System.out.println("           // Authored By swapnil.yavalkar //             ");
		System.out.println("======================================================================");
		System.out.println("Initiating CMS Connection.....");
		ISessionMgr sm = CrystalEnterprise.getSessionMgr();
		es = sm.logon(userId, password, CMS, "secEnterprise");
		iStore = (IInfoStore) es.getService("", "InfoStore");
	}

	public void setPassword(String s1) {
		this.password=s1;
	}

	public void setUserId(String s1){
	this.userId = s1;
	}

	public void setCMS(String s1) {
	this.CMS = s1;
	}

	public void setUniverseID(String s1) {
		this.universeid = s1;
		}
	
	public void setUniverseName(String s1) {
		this.universename = s1;
		}

	public void initExcel() {
			System.out.println("Initiating Excel.....");
			
			workbook = new XSSFWorkbook();
			sheetnotes = workbook.createSheet("Notes");
			sheet = workbook.createSheet("ReportDetails");
			my_style_1 = workbook.createCellStyle();
			my_style_2 = workbook.createCellStyle();
			XSSFFont font= workbook.createFont();
		    font.setFontHeightInPoints((short)10);
		    font.setFontName("Arial");
		    font.setColor(IndexedColors.WHITE.getIndex());
		    font.setBold(true);
		    font.setItalic(false);
			
		    my_style_1.setBorderBottom(BorderStyle.THIN);
		    my_style_1.setBorderTop(BorderStyle.THIN);
		    my_style_1.setBorderRight(BorderStyle.THIN);
		    my_style_1.setBorderLeft(BorderStyle.THIN);
		    
		    my_style_1.setAlignment(HorizontalAlignment.CENTER);
		    my_style_1.setVerticalAlignment(VerticalAlignment.CENTER);
		    
		    my_style_2.setFillForegroundColor(IndexedColors.ROYAL_BLUE.getIndex());
		    my_style_2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		    my_style_2.setFont(font);
		    
		    my_style_2.setBorderBottom(BorderStyle.THIN);
		    my_style_2.setBorderTop(BorderStyle.THIN);
		    my_style_2.setBorderRight(BorderStyle.THIN);
		    my_style_2.setBorderLeft(BorderStyle.THIN);
		    
		    my_style_2.setAlignment(HorizontalAlignment.CENTER);
		    my_style_2.setVerticalAlignment(VerticalAlignment.CENTER);
		    addNRow("======================================================================");
			addNRow("           // Authored By swapnil.yavalkar //             ");
			addNRow("======================================================================");
			System.out.println("Fetching the Report Details from the CMS database.....");
	}

	public void saveExcel() throws Exception{
	try {
		filepath = System.getProperty("user.dir");
		String finalpath = filepath+"\\"+finalfilename;
		FileOutputStream out = new FileOutputStream(new File(finalpath));
		workbook.write(out);
		out.close();
		System.out.println("Excel written successfully At: " + finalpath);
		
		} catch (Exception e) {
		e.printStackTrace();
		}
	}

	public void addNRow(String rowData) {
		int cellnum = 0;
		Row rown = sheetnotes.createRow(rownumn++);
		StringTokenizer st = new StringTokenizer(rowData, ":");
		while (st.hasMoreTokens()) {
			sheetnotes.autoSizeColumn(cellnum);
			Cell cell = rown.createCell(cellnum++);
			cell.setCellStyle(my_style_2);
			cell.setCellValue(st.nextToken());
		}		
	}
	
	public void addRow(String rowData) {
		int cellnum = 0;
		Row row = sheet.createRow(rownum++);
		StringTokenizer st = new StringTokenizer(rowData, "|");
			while (st.hasMoreTokens()) {
				if (row.getRowNum()==0)
				{
					sheet.autoSizeColumn(cellnum);
					Cell cell = row.createCell(cellnum++);
					cell.setCellStyle(my_style_2);
					sheet.setAutoFilter(CellRangeAddress.valueOf("A1:F1"));
					cell.setCellValue(st.nextToken());
				}else
			
				{
					sheet.autoSizeColumn(cellnum);
					Cell cell = row.createCell(cellnum++);
					cell.setCellStyle(my_style_1);
					cell.setCellValue(st.nextToken());
				}
			}
		}
	
	public void setFilename() {
		
		//finalfilename = CMS+"_"+universename+"_"+universeid+".xlsx";
		finalfilename = "ListOfUniverseBasedReports"+".xlsx";
	}
	
}