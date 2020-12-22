package ReadAndWrite;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class readExcel {

@Test (priority = 2)
public void readExcel()throws IOException{

        File file=new File("/Users/hakanozcan/Desktop/loginData.xlsx");
        FileInputStream fIP=new FileInputStream(file);
        XSSFWorkbook workbook=new XSSFWorkbook(fIP);

        if(file.isFile()&&file.exists()){

        XSSFSheet sheet=workbook.getSheetAt(0);

        Row rowName=sheet.getRow(0);
        Cell cellName=rowName.getCell(0);
        System.out.println(sheet.getRow(0).getCell(0));

        Row rowSurname=sheet.getRow(0);
        Cell cellSurname=rowSurname.getCell(1);
        System.out.println(sheet.getRow(0).getCell(1));

        Row rowEmail=sheet.getRow(0);
        Cell cellEmail=rowEmail.getCell(2);
        System.out.println(sheet.getRow(0).getCell(2));

        Row rowNameTest=sheet.getRow(1);
        Cell cellNameTest=rowNameTest.getCell(0);
        System.out.println(sheet.getRow(1).getCell(0));

        Row rowSurnameTest=sheet.getRow(1);
        Cell cellSurnameTest=rowSurnameTest.getCell(1);
        System.out.println(sheet.getRow(1).getCell(1));

        Row rowEmailTest=sheet.getRow(1);
        Cell cellEmailTest=rowEmailTest.getCell(2);
        System.out.println(sheet.getRow(1).getCell(2));

        }
        else
        {
        System.out.println("Excel data was not found.");
        }

        }
}
