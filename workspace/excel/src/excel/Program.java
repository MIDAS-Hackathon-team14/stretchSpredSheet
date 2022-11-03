package excel;


import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Chart;
import com.aspose.cells.ChartCollection;
import com.aspose.cells.ChartType;
import com.aspose.cells.SeriesCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

public class Program{
	public static void main(String[] args) throws Exception {
		Workbook workbook1 = new Workbook("C:\\web_1900_ljh\\시트 연습\\개인근무현황 시트.xlsx"); 

		// 첫 번째 워크시트의 참조 가져오기
		WorksheetCollection worksheets = workbook1.getWorksheets();
		Worksheet sheet = worksheets.get(0);

		// 셀에 일부 샘플 값 추가
		Cells cells = sheet.getCells();
		Cell cell = cells.get("A1");
		cell.setValue(50);
		cell = cells.get("A2");
		cell.setValue(100);
		cell = cells.get("A3");
		cell.setValue(150);
		cell = cells.get("B1");
		cell.setValue(4);
		cell = cells.get("B2");
		cell.setValue(20);
		cell = cells.get("B3");
		cell.setValue(50);

		// 워크시트에서 차트 가져오기
		ChartCollection charts = sheet.getCharts();

		// 워크시트에 차트 추가
		int chartIndex = charts.add(ChartType.PYRAMID, 5, 0, 15, 5);
		Chart chart = charts.get(chartIndex);

		// "A1" 범위의 차트에 NSeries(차트 데이터 소스) 추가
		// 셀을 "B3"으로
		SeriesCollection serieses = chart.getNSeries();
		serieses.add("A1:B3", true);
		
		workbook1.save("Excel_with_Chart.xlsx");
	}
}