package excel;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.examples.xssf.usermodel.BarChart;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAxDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBoolean;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTCatAx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLegend;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumRef;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTScaling;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrRef;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTValAx;
import org.openxmlformats.schemas.drawingml.x2006.chart.STAxPos;
import org.openxmlformats.schemas.drawingml.x2006.chart.STBarDir;
import org.openxmlformats.schemas.drawingml.x2006.chart.STCrosses;
import org.openxmlformats.schemas.drawingml.x2006.chart.STLegendPos;
import org.openxmlformats.schemas.drawingml.x2006.chart.STOrientation;
import org.openxmlformats.schemas.drawingml.x2006.chart.STTickLblPos;

public class UserWork {
	public static void main(String[] args) throws FileNotFoundException, IOException{
		String name="이준호";
		String userWorkStart[]="08:17:45 07:19:42 09:17:45 08:00:00 09:01:11 10:00:17 7:40:13".split(" "); //출근시간
    	String userWorkEnd[]="16:17:50 15:20:20 18:27:32 17:38:49 18:48:58 18:37:20 15:50:14".split(" "); //퇴근시간
    	String workWeek[]="2022-10-31 2022-11-01 2022-11-02 2022-11-03 2022-11-04 2022-11-05 2022-11-06".split(" "); //근무요일
		int userWorkHour[] = {8,8,9,9,9,8,8};//근무시간
		
		String day[] = "월 화 수 목 금 토 일".split(" ");
		String column[] = "근무날짜 요일 출근시간 퇴근시간 근무시간".split(" ");
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet(name+"사원 1주일 근무현황");
		XSSFRow header;
		sheet.setDisplayGridlines(false);
		Cell cell=null;
		sheet.addMergedRegion(new CellRangeAddress(1, 1	, 1,6));
		sheet.addMergedRegion(new CellRangeAddress(11, 11, 1,7));
		CellStyle style = workbook.createCellStyle();
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		
		header= sheet.createRow(1);
		cell = header.createCell(1);
		cell.setCellValue(name+"사원의 근무현황");
		
		header= sheet.createRow(11);
		cell = header.createCell(1);
		cell.setCellValue(name+"사원의 1주간 근무 통계");
		header.setHeight((short)500);
		CellStyle style2 = workbook.createCellStyle();
		style2.setVerticalAlignment(VerticalAlignment.CENTER);
		cell.setCellStyle(style2);
		
		header= sheet.createRow(2);	
		
		CellStyle style3 = workbook.createCellStyle();
		style3.setAlignment(HorizontalAlignment.RIGHT);
		style3.setVerticalAlignment(VerticalAlignment.CENTER);
		
		style3.setBorderBottom(BorderStyle.THIN);
		style3.setBorderTop(BorderStyle.THIN);
		style3.setBorderLeft(BorderStyle.THIN);
		style3.setBorderRight(BorderStyle.THIN);
		
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		
		for(int i = 0; i < 5; i++) {
			cell = header.createCell(i+1);
			cell.setCellValue(column[i]);
			cell.setCellStyle(style);
			sheet.setColumnWidth(i+1, 5000);
			header.setHeight((short)500);
		}
		
		for(int i = 0; i < userWorkStart.length; i++) {
			header= sheet.createRow(i+3);
			cell = header.createCell(1);
			cell.setCellValue(workWeek[i]);			
			cell.setCellStyle(style);
			
			cell = header.createCell(2);
			cell.setCellValue(day[i]);
			cell.setCellStyle(style);
			
			cell = header.createCell(3);
			cell.setCellValue(userWorkStart[i]);
			cell.setCellStyle(style);
			
			cell = header.createCell(4);
			cell.setCellValue(userWorkEnd[i]);
			cell.setCellStyle(style);
			
			cell = header.createCell(5);
			cell.setCellValue(userWorkHour[i]);
			cell.setCellStyle(style3);
		
			
		}
//		개인 회원 시트
//		차트 시작
		
//		XSSFSheet sheet2 = workbook.createSheet(name+"사원 요일별 근무시간 통계");
//		sheet2.setDisplayGridlines(false);
		
		XSSFDrawing drawing = (XSSFDrawing)sheet.createDrawingPatriarch();
		ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 12, 15, 22);

		XSSFChart chart = drawing.createChart(anchor);

		CTChart ctChart = ((XSSFChart)chart).getCTChart();
		CTPlotArea ctPlotArea = ctChart.getPlotArea();

		//여기까지는 똑같음

		//the first bar chart
		CTBarChart ctBarChart = ctPlotArea.addNewBarChart();
		CTBoolean ctBoolean = ctBarChart.addNewVaryColors();
		ctBoolean.setVal(true);
		ctBarChart.addNewBarDir().setVal(STBarDir.COL);

		//the first chart series
		CTBarSer ctBarSer = ctBarChart.addNewSer();
		CTSerTx ctSerTx = ctBarSer.addNewTx();
		CTStrRef ctStrRef = ctSerTx.addNewStrRef();
		ctStrRef.setF("'"+name+"사원 1주일 근무현황'!$C$4:$C10");
		ctBarSer.addNewIdx().setVal(0);
		CTAxDataSource ctAxDataSource = ctBarSer.addNewCat();
		ctStrRef = ctAxDataSource.addNewStrRef();
		ctStrRef.setF("'"+name+"사원 1주일 근무현황'!$C$4:$C10");
		CTNumDataSource ctNumDataSource = ctBarSer.addNewVal();
		CTNumRef ctNumRef = ctNumDataSource.addNewNumRef();
		ctNumRef.setF("'"+name+"사원 1주일 근무현황'!$C$4:$C10");

		//the second chart series
		CTBarSer ctBarSer1 = ctBarChart.addNewSer();
		CTSerTx ctSerTx1 = ctBarSer1.addNewTx();
		CTStrRef ctStrRef1 = ctSerTx1.addNewStrRef();
		ctStrRef1.setF("'"+name+"사원 1주일 근무현황'!$F$4:$F10");
		ctBarSer1.addNewIdx().setVal(1);
		CTAxDataSource ctAxDataSource1 = ctBarSer1.addNewCat();
		ctStrRef1 = ctAxDataSource1.addNewStrRef();
		ctStrRef1.setF("'"+name+"사원 1주일 근무현황'!$F$4:$F10");
		CTNumDataSource ctNumDataSource1 = ctBarSer1.addNewVal();
		CTNumRef ctNumRef1 = ctNumDataSource1.addNewNumRef();
		ctNumRef1.setF("'"+name+"사원 1주일 근무현황'!$F$4:$F10");
		
		// 3번째

		//at least the border lines in Libreoffice Calc ;-)
		ctBarSer.addNewSpPr().addNewLn().addNewSolidFill().addNewSrgbClr().setVal(new byte[] {0,0,0});
		ctBarSer1.addNewSpPr().addNewLn().addNewSolidFill().addNewSrgbClr().setVal(new byte[] {0,0,0});

		ctBarChart.addNewAxId().setVal(123456); //cat axis 1 (lines)
		ctBarChart.addNewAxId().setVal(123457); //val axis 1 (left)

		//cat axis 1
		CTCatAx ctCatAx = ctPlotArea.addNewCatAx();
		ctCatAx.addNewAxId().setVal(123456); //id of the cat axis
		CTScaling ctScaling = ctCatAx.addNewScaling();
		ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
		ctCatAx.addNewDelete().setVal(false);
		ctCatAx.addNewAxPos().setVal(STAxPos.B);
		ctCatAx.addNewCrossAx().setVal(123457); //id of the val axis
		ctCatAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);

		//val axis 1 (left)
		CTValAx ctValAx = ctPlotArea.addNewValAx();
		ctValAx.addNewAxId().setVal(123457); //id of the val axis
		ctScaling = ctValAx.addNewScaling();
		ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
		ctValAx.addNewDelete().setVal(false);
		ctValAx.addNewAxPos().setVal(STAxPos.L);
		ctValAx.addNewCrossAx().setVal(123456); //id of the cat axis
		ctValAx.addNewCrosses().setVal(STCrosses.AUTO_ZERO); //this val axis crosses the cat axis at zero
		ctValAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);
		
		File f = new File("C:\\web_1900_ljh\\시트 연습\\개인근무현황 시트.xlsx");
		if(f.exists()) {
		} else {
		      FileOutputStream fos= new FileOutputStream("C:\\web_1900_ljh\\시트 연습\\개인근무현황 시트.xlsx");	
		      workbook.write(fos);
		      fos.close();
		}
	}
}
