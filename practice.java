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
                String[] finalHeaders = null;
                for (int i = 0; i < sourceWorkbook.getNumberOfSheets(); i++) {

                    Row row4 = sourceSheet.getRow(3);
                    Cell cellE4 = row4.getCell(4);
                    String cellvalue = cellE4.getStringCellValue();
                    System.out.println(cellvalue);
                    String prefix = cellvalue.split(" ")[0];
                    System.out.println(prefix);
                    String[] dynamicheaders = {prefix + "_Content", prefix + "_Link"};

                    finalHeaders = Stream.concat(Arrays.stream(headers), Arrays.stream(dynamicheaders))
                            .toArray(String[]::new);

                    System.out.println(finalHeaders);

                }

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

                AtomicInteger index = new AtomicInteger(0);
                Arrays.stream(finalHeaders).forEachOrdered(header ->
                {
                    Cell cell = row.createCell(index.getAndIncrement());
                    cell.setCellValue(header);
                });

                Row newexcetrow = sheet.getRow(0);
                int masterModulesColumnIndex = -1;
                int sgEnContentColumnIndex = -1;
                int masterElementsColumnIndex = -1;

                // Find "Master_Modules", "Master_Elements", and "SG_EN Content" column indices
                for (Cell cell : newexcetrow) {
                    if (cell.getStringCellValue().equalsIgnoreCase("Master_Modules")) {
                        masterModulesColumnIndex = cell.getColumnIndex();
                    } else if (cell.getStringCellValue().equalsIgnoreCase("Master_Elements")) {
                        masterElementsColumnIndex = cell.getColumnIndex();
                    } else if (cell.getStringCellValue().equalsIgnoreCase("SG_EN Content")) {
                        sgEnContentColumnIndex = cell.getColumnIndex();
                    }
                }

                if (masterModulesColumnIndex == -1) {
                    System.out.println("Master_Modules column not found in the new sheet");
                } else {
                    int newrowindex = 1;
                    for (String moduleinput : moduleinputs) {
                        Row newrow = sheet.createRow(newrowindex);
                        Cell newCell = newrow.getCell(masterModulesColumnIndex);
                        if (newCell == null) {
                            newCell = newrow.createCell(masterModulesColumnIndex);
                        }
                        newCell.setCellValue(moduleinput);
                        newrowindex++;
                    }

                    // Requirement 1: Add rows below "Preheader" in Master_Elements
                    int preheaderRowIndex = -1;
                    for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                        Row currentRow = sheet.getRow(rowIndex);
                        if (currentRow != null) {
                            Cell cell = currentRow.getCell(masterModulesColumnIndex);
                            if (cell != null && cell.getStringCellValue().equalsIgnoreCase("Preheader")) {
                                preheaderRowIndex = rowIndex;
                                Cell preheaderCell = currentRow.createCell(masterElementsColumnIndex);
                                preheaderCell.setCellValue("ps");
                                break;
                            }
                        }
                    }

                    if (preheaderRowIndex != -1) {
                        // Insert "ssl" and "vo" rows below the "Preheader" row
                        Row sslRow = sheet.createRow(preheaderRowIndex + 1);
                        Cell sslModuleCell = sslRow.createCell(masterModulesColumnIndex);
                        sslModuleCell.setCellValue("ssl");
                        Cell sslElementCell = sslRow.createCell(masterElementsColumnIndex);
                        sslElementCell.setCellValue("ssl");

                        Row voRow = sheet.createRow(preheaderRowIndex + 2);
                        Cell voModuleCell = voRow.createCell(masterModulesColumnIndex);
                        voModuleCell.setCellValue("vo");
                        Cell voElementCell = voRow.createCell(masterElementsColumnIndex);
                        voElementCell.setCellValue("vo");
                    }

                    // Requirement 2: Fetch the value from column E for "Preheader" in EMAIL 1 and update SG_EN Content for "ssl"
                    String preheaderValueForSG_EN = null;

                    // Locate "Preheader" in source sheet and fetch value from column E (index 4)
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

                    // Add the fetched value to SG_EN Content column for the "ssl" row
                    if (preheaderValueForSG_EN != null) {
                        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                            Row currentRow = sheet.getRow(rowIndex);
                            if (currentRow != null) {
                                Cell cell = currentRow.getCell(masterModulesColumnIndex);
                                if (cell != null && cell.getStringCellValue().equalsIgnoreCase("ssl")) {
                                    Cell sgEnCell = currentRow.createCell(sgEnContentColumnIndex);
                                    sgEnCell.setCellValue(preheaderValueForSG_EN);
                                    break;
                                }
                            }
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
