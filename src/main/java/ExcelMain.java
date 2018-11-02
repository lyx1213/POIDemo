import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

public class ExcelMain {
    public static final String fileName = "MyFirstExcel.xlsx";
    public static final String fileName2 = "20180918.xlsx";
    public static final String fileName3 = "20180919.xlsx";




    public static void test(){
        try {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheetname = workbook.createSheet("DataTypes");
            Object[][] dataTypes = {
                    {"name", "address"},
                    {"lyx", "1001"},
                    {"cc", "sad"}
            };
            int rowNum = 0;

            for (Object[] dataType : dataTypes) {
                XSSFRow row = sheetname.createRow(rowNum++);
                int colNum = 0;
                for (Object filed : dataType) {
                    XSSFCell cell = row.createCell(colNum++);
                    cell.setCellValue(filed.toString());
                }
            }

            FileOutputStream outputStream = new FileOutputStream(fileName);
            workbook.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    public static List<String[]> readyData(){
        List<String[]> shopList = new ArrayList<String[]>();
        String[] rowHeadData = {"shopName", "hour","count"};
        shopList.add(rowHeadData);
        try {
            String url = "/Users/jack/Documents/wina/customer/20180919";

            File file = new File(url);
            File[] shopFiles = file.listFiles();
            for(int i=0;i<shopFiles.length;i++){

                String shopName = shopFiles[i].getName();
                //delete part.gz
                if(shopFiles[i].isDirectory()){
                    File[] hourFiles = shopFiles[i].listFiles();
                    for(int j=0;j<hourFiles.length;j++){
                        String hour = hourFiles[j].getName();
                        if(hourFiles[j].isDirectory()){
                            File[] files1 = hourFiles[j].listFiles();
                            for (File file1 : files1) {
                                String absolutePath = file1.getAbsolutePath();
                                int totalLines = getTotalLines(absolutePath);
                                String[] rowData = {shopName,hour,String.valueOf(totalLines)};
                                shopList.add(rowData);
                            }
                        }
                    }
                }

            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return shopList;

    }

    public static int getTotalLines(String path){
        int total = 0;
        try {
            List<String> stringList = Files.readAllLines(Paths.get(path));
            total = stringList.size();

        } catch (IOException e) {
            e.printStackTrace();
        }
        return total;
    }

    public static void createShopData(){
        try {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheetname = workbook.createSheet("shopData");
            List<String[]> shopList = readyData();

            for (int i=0;i<shopList.size();i++){
                XSSFRow row = sheetname.createRow(i);
                for(int j = 0;j<3;j++){
                    XSSFCell cell = row.createCell(j);
                    cell.setCellValue(shopList.get(i)[j]);

                }
            }
            FileOutputStream outputStream = new FileOutputStream(fileName3);
            workbook.write(outputStream);
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }


    }

    public static void main(String[] args){
        createShopData();
        //readyData();
    }

}
