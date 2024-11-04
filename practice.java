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
        String[] headers = {"Master_Modules", "Module_Background Color", "Master_Elements", "SG_EN Content"};

        try (FileInputStream fis = new FileInputStream(sourceFilePath)) {
            Workbook sourceWorkbook = new XSSFWorkbook(fis);
            Sheet sourceSheet = sourceWorkbook.getSheet("EMAIL 1");

            // Write headers to the new sheet
            AtomicInteger index = new AtomicInteger(0);
            Arrays.stream(headers).forEachOrdered(header -> {
                Cell cell = row.createCell(index.getAndIncrement());
                cell.setCellValue(header);
            });

            // Find the "MODULES" column index in the source sheet
            int modulesColumnIndex = -1;
            Row sourceHeaderRow = sourceSheet.getRow(4); // Assumes headers are in row 4

            for (Cell cell : sourceHeaderRow) {
                if (cell.getStringCellValue().equalsIgnoreCase("MODULES")) {
                    modulesColumnIndex = cell.getColumnIndex();
                    break;
                }
            }

            if (modulesColumnIndex == -1) {
                System.out.println("There is no column called 'MODULES' in the input file.");
                return;
            }

            // Read module data from source sheet and add it to the new sheet
            List<String> moduleInputs = new ArrayList<>();
            for (int i = 5; i <= sourceSheet.getLastRowNum(); i++) {  // Start from row 5, assuming data starts below headers
                Row sourceRow = sourceSheet.getRow(i);
                if (sourceRow != null) {
                    Cell moduleCell = sourceRow.getCell(modulesColumnIndex);
                    if (moduleCell != null && moduleCell.getCellType() == Cell.CELL_TYPE_STRING) {
                        moduleInputs.add(moduleCell.getStringCellValue());
                    }
                }
            }

            int newRowIdx = 1; // Start adding data from the second row
            int masterModulesColIndex = 0;
            int masterElementsColIndex = 2;
            int sgEnContentColIndex = 3;

            for (String moduleInput : moduleInputs) {
                Row newRow = sheet.createRow(newRowIdx++);
                Cell moduleCell = newRow.createCell(masterModulesColIndex);
                moduleCell.setCellValue(moduleInput);

                // If the module is "Preheader," add ps, ssl, and vo rows and fetch the "SG_EN Content" value
                if (moduleInput.equalsIgnoreCase("Preheader")) {
                    String sgEnContentValue = null;

                    // Find and fetch the SG_EN Content from column E of the "Preheader" row in the source sheet
                    for (int i = 5; i <= sourceSheet.getLastRowNum(); i++) {
                        Row sourceRow = sourceSheet.getRow(i);
                        if (sourceRow != null) {
                            Cell moduleCellCheck = sourceRow.getCell(modulesColumnIndex);
                            if (moduleCellCheck != null && moduleCellCheck.getStringCellValue().equalsIgnoreCase("Preheader")) {
                                Cell sgEnCell = sourceRow.getCell(4); // Column E index is 4
                                if (sgEnCell != null) {
                                    sgEnContentValue = sgEnCell.getStringCellValue();
                                }
                                break;
                            }
                        }
                    }

                    // Add "ps" in the same row as "Preheader"
                    Cell elementCell = newRow.createCell(masterElementsColIndex);
                    elementCell.setCellValue("ps");

                    // Add additional rows for "ssl" and "vo"
                    Row sslRow = sheet.createRow(newRowIdx++);
                    sslRow.createCell(masterModulesColIndex).setCellValue("Preheader");
                    sslRow.createCell(masterElementsColIndex).setCellValue("ssl");
                    if (sgEnContentValue != null) {
                        sslRow.createCell(sgEnContentColIndex).setCellValue(sgEnContentValue); // Set SG_EN Content for ssl
                    }

                    Row voRow = sheet.createRow(newRowIdx++);
                    voRow.createCell(masterModulesColIndex).setCellValue("Preheader");
                    voRow.createCell(masterElementsColIndex).setCellValue("vo");
                }
            }

            // Write the workbook to a file
            try (FileOutputStream fos = new FileOutputStream("ProvarExcel.xlsx")) {
                workbook.write(fos);
                System.out.println("Excel file created successfully.");
            } catch (IOException e) {
                e.printStackTrace();
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
