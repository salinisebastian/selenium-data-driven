package salini.company;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Data {
	private int noOfColumns = 0;

	public int getnoOfColumns() {
		return noOfColumns;
	}

	public ArrayList<String> getData(String ClassDivision, String Sheetname) throws IOException {
		FileInputStream DataFile = new FileInputStream("C:\\Users\\PC\\Downloads\\data.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(DataFile);

		ArrayList<String> array = new ArrayList<String>();
		int CountOfSheets = workbook.getNumberOfSheets();

		for (int i = 0; i < CountOfSheets; i++) {
			if (workbook.getSheetName(i).equalsIgnoreCase(Sheetname)) {

				XSSFSheet sheet = workbook.getSheetAt(i);
				Iterator<Row> row = sheet.iterator();

				Row firstrow = row.next();

				Iterator<Cell> cell = firstrow.iterator();
				int j = 0;
				int col = 0;
				while (cell.hasNext()) {
					if (cell.next().getStringCellValue().equalsIgnoreCase("Classes")) {
						col = j;
					}

					j++;

				}

				while (row.hasNext()) {
					Row requiredrow = row.next();
					// System.out.println(requiredrow.getCell(col).getStringCellValue());
					noOfColumns = requiredrow.getPhysicalNumberOfCells();
					if (requiredrow.getCell(col).getStringCellValue().equalsIgnoreCase(ClassDivision))

					{
						Iterator<Cell> c = requiredrow.cellIterator();
						while (c.hasNext()) {
							Cell o = c.next();
							if (o.getCellType() == CellType.STRING) {
								array.add(o.getStringCellValue());
							} else {
								array.add(NumberToTextConverter.toText(o.getNumericCellValue()));
							}
						}

						break;
					}

				}
			}

		}
		return array;
	}

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
	}
}
