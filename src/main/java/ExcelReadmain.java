import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

public class ExcelReadmain {

    public static void mergeCell(){
        try {
            String FILE_NAME = "/Users/jack/Documents/wina/temp2/20180919.xlsx";

            FileInputStream inputStream = new FileInputStream(new File(FILE_NAME));

            Workbook workbook = new XSSFWorkbook(inputStream);

            Sheet sheetAt = workbook.getSheetAt(0);
            int firstRow = 1;
            for(int i =1;i<365;i++){
                sheetAt.addMergedRegion(new CellRangeAddress(firstRow,firstRow+13,0,0));
                firstRow = firstRow+14;

            }

            //sheetAt.addMergedRegion(new CellRangeAddress(1,14,0,0));
            Workbook workbook1 = sheetAt.getWorkbook();
            FileOutputStream outputStream = new FileOutputStream("20180919_2.xlsx");
            workbook1.write(outputStream);
            workbook1.close();

        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    public static void test(){
        try {
            String FILE_NAME = System.getProperty("user.home").concat("/Documents/公司汇总18.xlsx");

            FileInputStream inputStream = new FileInputStream(new File(FILE_NAME));

            Workbook workbook = new XSSFWorkbook(inputStream);

            Sheet sheetAt = workbook.getSheetAt(0);


            Workbook workbook1 = sheetAt.getWorkbook();
            FileOutputStream outputStream = new FileOutputStream("myStrExcel.xlsx");
            workbook1.write(outputStream);
            outputStream.close();
            inputStream.close();


/*        Iterator<Row> iterator = sheetAt.iterator();
        int num = 0;
        while (iterator.hasNext()&&num<=300){
            num++;
            Row next = iterator.next();
            Cell cell = next.getCell(1);
            if(cell==null){
                continue;
            }
            if(cell.getCellTypeEnum() == CellType.STRING){
                System.out.println(cell.getStringCellValue());
            }
            if(cell.getCellTypeEnum() == CellType.NUMERIC){
                System.out.println(cell.getNumericCellValue());
            }

        }*/
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void getShooNum(){

        String url = "/Users/jack/Documents/wina/customer/20180918";

        File file = new File(url);
        File[] shopFiles = file.listFiles();
        System.out.println(shopFiles.length);
        /*for(int i=0;i<shopFiles.length;i++){

            String shopCode = shopFiles[i].getName();
            //delete part.gz

        }*/

    }

    public static void main(String[] args) {

        //getShooNum();
        mergeCell();

    }

}
