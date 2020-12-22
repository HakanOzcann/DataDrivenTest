package ReadAndWrite;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class writeExcel {

    @Test(priority = 1)
    public void ExcelReadTest() throws IOException {

        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sh = wb.createSheet("TestSheet");

        XSSFRow row;
        Map<String, Object[]> loginInfo = new TreeMap<String, Object[]>();

        loginInfo.put("1", new Object[]{
                "Hakan", "Ozcan", "hakan.ozcan44@hotmail.com"});
        loginInfo.put("2", new Object[]{
                "TestName", "TestSurname", "TestEmail"});

        Set<String> keyId = loginInfo.keySet();
        int rowId = 0;

        for (String key : keyId) {
            row = sh.createRow(rowId++);
            Object[] objectArr = loginInfo.get(key);
            int cellId = 0;

            for (Object obj : objectArr) {
                Cell cell = row.createCell(cellId++);
                cell.setCellValue((String) obj);
            }
        }
        String fpath = "/Users/hakanozcan/Desktop/loginData.xlsx";
        FileOutputStream out = new FileOutputStream(new File(fpath));
        wb.write(out);
        out.close();
        System.out.println("loginData.xlsx written successfully");
    }

}



