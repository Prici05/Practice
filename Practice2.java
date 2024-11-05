// Assuming masterElementsColumnIndex is the index for "Master_Elements"
// Find "SG-EN_Content" column index in the output sheet
int sgEnContentColumnIndex = -1;
for (Cell cell : newexcetrow) {
    if (cell.getStringCellValue().equalsIgnoreCase("SG-EN_Content")) {
        sgEnContentColumnIndex = cell.getColumnIndex();
        break;
    }
}

// Now, fetch the value from source sheet
if (sgEnContentColumnIndex != -1) {
    for (int rowIndex = 1; rowIndex <= sourceSheet.getLastRowNum(); rowIndex++) {
        Row sourceRow = sourceSheet.getRow(rowIndex);
        if (sourceRow != null) {
            Cell moduleCell = sourceRow.getCell(reqcolindex);
            if (moduleCell != null && moduleCell.getStringCellValue().equalsIgnoreCase("Preheader")) {
                // Fetch the value from cell E (column index 4)
                Cell valueCell = sourceRow.getCell(4);
                String fetchedValue = valueCell != null ? valueCell.getStringCellValue() : "";

                // Now paste that value in the output sheet corresponding to the "ssl" row
                Row sslRow = sheet.getRow(rowIndex + 1); // Assuming ssl row is directly after Preheader row
                if (sslRow != null) {
                    Cell sgEnContentCell = sslRow.createCell(sgEnContentColumnIndex);
                    sgEnContentCell.setCellValue(fetchedValue);
                }
                break; // Exit after processing
            }
        }
    }
}
