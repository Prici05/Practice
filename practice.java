import java.util.Arrays;
import java.util.stream.Stream;
import java.io.FileOutputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ProvarExcel {

    public static void main(String[] args) throws FileNotFoundException, IOException {

        String sourceFilePath = "Program Brief Send to Receive - Asset.xlsx"; // Path to source file

        // Create a new workbook
        Workbook workbook = new XSSFWorkbook();

        // Create a new sheet
        Sheet sheet = workbook.createSheet("Sheet1");

        Row row = sheet.createRow(0);
        String[] headers = {"Master_Modules", "Module_Background Color", "Master_Elements"};

        try (FileInputStream fis = new FileInputStream(sourceFilePath)) {
            Workbook sourceWorkbook = new XSSFWorkbook(fis);

            Sheet sourceSheet = sourceWorkbook.getSheet("EMAIL 1");
            String[] finalHeaders = null;

            // Dynamically get headers
            for (int i = 0; i < sourceWorkbook.getNumberOfSheets(); i++) {
                Row row4 = sourceSheet.getRow(3);
                Cell cellE4 = row4.getCell(4);
                String cellvalue = cellE4.getStringCellValue();
                System.out.println(cellvalue);

                String prefix = cellvalue.split("")[0];
                System.out.println(prefix);

                String[] dynamicheaders = {prefix + "_Content", prefix + "_Link"};
                finalHeaders = Stream.concat(Arrays.stream(headers), Arrays.stream(dynamicheaders)).toArray(String[]::new);

                System.out.println(Arrays.toString(finalHeaders));

                Row headingrow = sourceSheet.getRow(4);
                int reqcolindex = -1;

                // Find the "MODULES" column index
                for (Cell cell : headingrow) {
                    if (cell.getStringCellValue().equalsIgnoreCase("MODULES")) {
                        reqcolindex = cell.getColumnIndex();
                        break;
                    }
                }

                if (reqcolindex == -1) {
                    System.out.println("There is no column called Modules in the input file");
                } else {
                    List<String> moduleinputs = new ArrayList<>();
                    for (int j = 1; j <= sourceSheet.getLastRowNum(); j++) {  // Skip the header (row 0)
                        Row destRow = sourceSheet.getRow(j);
                        if (destRow != null) {
                            Cell destcell = destRow.getCell(reqcolindex);
                            if (destcell != null && destcell.getCellType() == Cell.CELL_TYPE_STRING) {
                                moduleinputs.add(destcell.getStringCellValue());
                            }
                        }
                    }

                    AtomicInteger index = new AtomicInteger(0);
                    Arrays.stream(finalHeaders).forEachOrdered(header -> {
                        Cell cell = row.createCell(index.getAndIncrement());
                        cell.setCellValue(header);
                    });

                    Row newexcetrow = sheet.getRow(0);
                    int masterModulesColumnIndex = -1;

                    // Find "Master_Modules" column index
                    for (Cell cell : newexcetrow) {
                        if (cell.getStringCellValue().equalsIgnoreCase("Master_Modules")) {
                            masterModulesColumnIndex = cell.getColumnIndex();
                            break;
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

                        // Now we handle the Preheader case and insert new rows
                        int masterElementsColumnIndex = -1;

                        // Find "Master_Elements" column index
                        for (Cell cell : newexcetrow) {
                            if (cell.getStringCellValue().equalsIgnoreCase("Master_Elements")) {
                                masterElementsColumnIndex = cell.getColumnIndex();
                                break;
                            }
                        }

                        if (masterModulesColumnIndex != -1 && masterElementsColumnIndex != -1) {
                            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                                Row currentRow = sheet.getRow(rowIndex);
                                if (currentRow != null) {
                                    Cell cell = currentRow.getCell(masterModulesColumnIndex);
                                    if (cell != null && cell.getStringCellValue().equalsIgnoreCase("Preheader")) {
                                        // Add "ps" in the same row as "Preheader"
                                        Cell newElementCell = currentRow.getCell(masterElementsColumnIndex);
                                        if (newElementCell == null) {
                                            newElementCell = currentRow.createCell(masterElementsColumnIndex);
                                        }
                                        newElementCell.setCellValue("ps");

                                        // Insert new rows for "ssl" and "vo"
                                        String[] newValues = {"ssl", "vo"};
                                        int insertRowIndex = rowIndex + 1; // Start inserting after the current row
                                        for (String newValue : newValues) {
                                            sheet.shiftRows(insertRowIndex, sheet.getLastRowNum(), 1); // Shift rows
                                            Row newRow = sheet.createRow(insertRowIndex);
                                            Cell newCell = newRow.createCell(masterElementsColumnIndex);
                                            newCell.setCellValue(newValue);
                                            insertRowIndex++;
                                        }
                                        break; // Exit after processing "Preheader"
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

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
