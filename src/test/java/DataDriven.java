import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriven {

	@SuppressWarnings("deprecation")
	public ArrayList<String> getData(String test_case_name) throws IOException
	{

		ArrayList<String> a = new ArrayList<String>();

		FileInputStream fs = new FileInputStream("D://DataDriven_Demo//DemoData.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fs);

		int sheet_count = workbook.getNumberOfSheets();
		for(int i = 0; i < sheet_count; i++)
		{
			if((workbook.getSheetName(i)).equalsIgnoreCase("Sheet1"))
			{

				XSSFSheet sheet = workbook.getSheetAt(i);
				Iterator<Row> rows = sheet.iterator();
				Row first_row = rows.next();
				Iterator<Cell> cell = first_row.cellIterator();
				int k = 0;
				int column = 0;

				while(cell.hasNext())
				{
					Cell cell_value = cell.next();

					if(cell_value.getStringCellValue().equalsIgnoreCase("Testcases"))
					{

						column = k;

					}

					k++;
				}
				while(rows.hasNext())
				{

					Row r = rows.next();
					if(r.getCell(column).getStringCellValue().equalsIgnoreCase(test_case_name))
					{
						Iterator<Cell> cv = r.cellIterator();

						while(cv.hasNext())
						{
							Cell c  = cv.next();
							if(c.getCellTypeEnum() == CellType.STRING)
							{
								a.add(cv.next().getStringCellValue());
							}
							else
							{
								a.add(NumberToTextConverter.toText(c.getNumericCellValue()));
							}
						}
					}

				}

			}

		}

		return a;

	}

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
	}

}
