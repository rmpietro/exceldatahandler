import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class Main {

    public static void main(String[] args) {
        Main main = new Main();
        System.out.println(main.readExcelContent("users.xlsx"));
    }

    public Map<Integer, Map<String, String>> readExcelContent(String resourceName) {
        try (FileInputStream file = new FileInputStream(new File(getClass().getClassLoader().getResource(resourceName).getFile()));
             Workbook workbook = new XSSFWorkbook(file);
            )
        {
            Sheet sheet = workbook.getSheetAt(0);
            Map<Integer, String> titles = new HashMap<>();

            Map<Integer, Map<String, String>> data = new HashMap<>();
            for (Row row : sheet) {
                int rowNum = row.getRowNum();
                Map<String, String> rowData = new HashMap<>();
                for (Cell cell : row) {
                    cell.setCellType(CellType.STRING);
                    if (rowNum == 0) {
                        String title = cell.getStringCellValue();
                        titles.put(cell.getColumnIndex(), cell.getStringCellValue());
                    } else {
                        rowData.put(titles.get(cell.getColumnIndex()), cell.getStringCellValue());
                    }
                }
                if (rowData.size() >= row.getLastCellNum())
                    data.put(rowNum, rowData);
            }
            return data;

        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }
}