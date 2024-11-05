package com.example;
import java.util.Arrays;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.stream.Stream;
import java.io.FileOutputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

public class ProvarExcel {

    public static void main(String[] args) throws FileNotFoundException, IOException {

        String sourceFilePath = "Program Brief - Send to Receive - Asset.xlsx"; // Path to source file

        // 1. create a new workbook
        Workbook workbook = new XSSFWorkbook();

        // 2. Create a new sheet
        Sheet sheet = workbook.createSheet("Sheet1");

        Row row = sheet.createRow(0);
        String[] headers = {"Master_Modules", "Module_Background_Color", "Master_Elements", "SG_EN Content"};

        try (FileInputStream fis = new FileInputStream(sourceFilePath)) {
            Workbook sourceWorkbook = new XSSFWorkbook(fis);
            {
                Sheet sourceSheet = sourceWorkbook.getSheet("EMAIL 1");
                String[] finalHeaders = headers;

                AtomicInteger index = new AtomicInteger(0);
                Arrays.stream(finalHeaders).forEachOrdered(header -> {
                    Cell cell = row.createCell(index.getAndIncrement());
                    cell.setCellValue(header);
                });

                Row headingrow = sourceSheet.getRow(4);
                int reqcolindex = -1;
                for (Cell cell : headingrow) {
                    if (cell.getStringCellValue().equalsIgnoreCase("MODULES")) {
                        reqcolindex = cell.getColumnIndex();
                        break;
                    }
                }
                if (reqcolindex == -1) {
                    System.out.println("There is no column called Modules in the input file excel");
                }

                List<String> moduleinputs = new ArrayList<>();
                for (int j = 5; j <= sourceSheet.getLastRowNum(); j++) {
                    Row destRow = sourceSheet.getRow(j);
                    if (destRow != null) {
                        Cell destcell = destRow.getCell(reqcolindex);
                        if (destcell != null && destcell.getCellType() == CellType.STRING) {
                            moduleinputs.add(destcell.getStringCellValue());
                        }
                    }
                }

                int masterModulesColumnIndex = 0; // Master_Modules is first in headers
                int masterElementsColumnIndex = 2; // Master_Elements is third in headers
                int sgEnContentColumnIndex = 3; // SG_EN Content is fourth in headers

                int newRowIdx = 1;
                for (String moduleinput : moduleinputs) {
                    Row newRow = sheet.createRow(newRowIdx++);
                    Cell moduleCell = newRow.createCell(masterModulesColumnIndex);
                    moduleCell.setCellValue(moduleinput);

                    if (moduleinput.equalsIgnoreCase("Preheader")) {
                        // Add `ps` in the same row as "Preheader"
                        Cell psCell = newRow.createCell(masterElementsColumnIndex);
                        psCell.setCellValue("ps");

                        // Add `ssl` and `vo` in the following rows
                        Row sslRow = sheet.createRow(newRowIdx++);
                        Cell sslModuleCell = sslRow.createCell(masterModulesColumnIndex);
                        sslModuleCell.setCellValue("ssl");
                        sslRow.createCell(masterElementsColumnIndex).setCellValue("ssl");

                        Row voRow = sheet.createRow(newRowIdx++);
                        Cell voModuleCell = voRow.createCell(masterModulesColumnIndex);
                        voModuleCell.setCellValue("vo");
                        voRow.createCell(masterElementsColumnIndex).setCellValue("vo");

                        // Retrieve value for SG_EN Content for "ssl" from EMAIL 1 sheet, column E
                        String preheaderValueForSG_EN = null;
                        for (int rowIndex = 1; rowIndex <= sourceSheet.getLastRowNum(); rowIndex++) {
                            Row currentRow = sourceSheet.getRow(rowIndex);
                            if (currentRow != null) {
                                Cell cell = currentRow.getCell(reqcolindex);
                                if (cell != null && cell.getStringCellValue().equalsIgnoreCase("Preheader")) {
                                    preheaderValueForSG_EN = currentRow.getCell(4).getStringCellValue();
                                    break;
                                }
                            }
                        }

                        if (preheaderValueForSG_EN != null) {
                            sslRow.createCell(sgEnContentColumnIndex).setCellValue(preheaderValueForSG_EN);
                        }
                    }
                }
            }
        }

        try (FileOutputStream fos = new FileOutputStream("ProvarExcel.xlsx")) {
            workbook.write(fos);
            System.out.println("Excel file created successfully");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
