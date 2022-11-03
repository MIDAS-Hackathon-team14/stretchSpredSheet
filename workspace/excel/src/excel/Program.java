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
		Workbook workbook1 = new Workbook("C:\\web_1900_ljh\\��Ʈ ����\\���αٹ���Ȳ ��Ʈ.xlsx"); 

		// ù ��° ��ũ��Ʈ�� ���� ��������
		WorksheetCollection worksheets = workbook1.getWorksheets();
		Worksheet sheet = worksheets.get(0);

		// ���� �Ϻ� ���� �� �߰�
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

		// ��ũ��Ʈ���� ��Ʈ ��������
		ChartCollection charts = sheet.getCharts();

		// ��ũ��Ʈ�� ��Ʈ �߰�
		int chartIndex = charts.add(ChartType.PYRAMID, 5, 0, 15, 5);
		Chart chart = charts.get(chartIndex);

		// "A1" ������ ��Ʈ�� NSeries(��Ʈ ������ �ҽ�) �߰�
		// ���� "B3"����
		SeriesCollection serieses = chart.getNSeries();
		serieses.add("A1:B3", true);
		
		workbook1.save("Excel_with_Chart.xlsx");
	}
}