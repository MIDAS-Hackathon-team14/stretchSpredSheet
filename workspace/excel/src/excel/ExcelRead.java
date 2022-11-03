package excel;
 
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
 
public class ExcelRead {
 
    public static void main(String[] args) throws FileNotFoundException, IOException{
    	String today = "2022-11-03"; //오늘날짜
    	String column[] = "사원명 출근시간 퇴근시간 근무시간".split(" ");
    	String userName[]="이준호 이동현 김은빈 조수현 이윤서 남주빈".split(" ");//유저
    	String userWorkStart[]="08:17:45 07:19:42 09:17:45 08:00:00 09:01:11 10:00:17".split(" ");//출근시간
    	String userWorkEnd[]="16:17:50 15:20:20 18:27:32 17:38:49 18:48:58 18:37:20".split(" ");//퇴근시간
    	int userWorkHour[]= {8,8,9,9,9,8};//근무시간
    	
		XSSFWorkbook workbook = new XSSFWorkbook();
		
		XSSFSheet sheet = workbook.createSheet(today+"일 근무현황");
		XSSFRow header;
		sheet.setDisplayGridlines(false);
		
		Cell cell=null;
		sheet.addMergedRegion(new CellRangeAddress(2, 2	, 3,6));
				
		CellStyle style = workbook.createCellStyle();
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		
		CellStyle style2 = workbook.createCellStyle();
		style2.setAlignment(HorizontalAlignment.CENTER);
		style2.setVerticalAlignment(VerticalAlignment.CENTER);
		
		CellStyle style3 = workbook.createCellStyle();
		style3.setAlignment(HorizontalAlignment.RIGHT);
		style3.setVerticalAlignment(VerticalAlignment.CENTER);
		
		header= sheet.createRow(2);
		cell = header.createCell(3);
		cell.setCellValue(today+"일의 근무현황");
		header.setHeight((short)500);
		cell.setCellStyle(style2);
		
		
		header= sheet.createRow(5);	
		
		style3.setBorderBottom(BorderStyle.THIN);
		style3.setBorderTop(BorderStyle.THIN);
		style3.setBorderLeft(BorderStyle.THIN);
		style3.setBorderRight(BorderStyle.THIN);
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		for(int i = 0; i < 4; i++) {
			cell = header.createCell(3+i);
			cell.setCellValue(column[i]);
			cell.setCellStyle(style);
			sheet.setColumnWidth(i+3, 5000);
			header.setHeight((short)500);
		}
				
		for(int i = 0; i < userName.length; i++) {
			header= sheet.createRow(6+i);
			cell = header.createCell(3);
			cell.setCellValue(userName[i]);
			cell.setCellStyle(style);
			cell = header.createCell(4);
			cell.setCellValue(userWorkStart[i]);
			cell.setCellStyle(style);
			cell = header.createCell(5);
			cell.setCellValue(userWorkEnd[i]);
			cell.setCellStyle(style);
			cell = header.createCell(6);
			cell.setCellValue(userWorkHour[i]);
			cell.setCellStyle(style3);
		}
		

		
		
		
		File f = new File("C:\\web_1900_ljh\\시트 연습\\전체사원.xlsx");
		if(f.exists()) {
		} else {
		      FileOutputStream fos= new FileOutputStream("C:\\web_1900_ljh\\시트 연습\\전체사원.xlsx");	
		      workbook.write(fos);
		      fos.close();
		}
    }
}	