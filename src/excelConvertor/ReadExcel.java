package excelConvertor;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.FileInputStream;

public class ReadExcel {

    public String getExcelData(String filePath, String sheetName, int[] cols) {
        StringBuilder excelData = new StringBuilder();
        try {
            FileInputStream fis = new FileInputStream(filePath);
            Workbook wb = WorkbookFactory.create(fis);
            Sheet sh = wb.getSheet(sheetName);
            if (sh == null)
                sh = wb.getSheetAt(0);

            int rows = sh.getLastRowNum();

            excelData.append("[\n");

            for (int i = 1; i <= rows; i++) {

                excelData.append("{\n");


                for (int j = cols[0]; j < cols.length; j++) {
                    excelData.append("\t\"");
                    excelData.append(sh.getRow(0).getCell(j));
                    excelData.append("\" : \"");
                    excelData.append(sh.getRow(i).getCell(j));
                    if (j < cols.length - 1)
                        excelData.append("\",\n ");
                    else
                        excelData.append("\"\n");
                }

                if (i < rows)
                    excelData.append("},\n");
                else
                    excelData.append("}\n");
            }
            excelData.append("]");
            System.out.println(rows + " rows converted successfully");

        } catch (Exception e) {
            e.printStackTrace();
        }
        return excelData.toString();
    }


}
