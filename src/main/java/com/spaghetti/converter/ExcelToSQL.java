package com.spaghetti.converter;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.StringUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedWriter;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDateTime;

public class ExcelToSQL {
    private static final Logger log = LogManager.getLogger(ExcelToSQL.class);

    public static void main(String[] args) {

        String excelFilePath = "file.xlsx"; //C:/Users/user/Downloads/file.xlsx
        String sqlFilePath = "file.sql";
        String tableName = "table_name";

        try (InputStream excelFile = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(excelFile);
             BufferedWriter sqlFile = Files.newBufferedWriter(Paths.get(sqlFilePath))) {

            Sheet sheet = workbook.getSheetAt(0);
            int rowCount = sheet.getPhysicalNumberOfRows();

            sqlFile.write("INSERT INTO " + tableName + " (");

            Row headerRow = sheet.getRow(0);
            int factColumnCount = 0;
            int columnCount = headerRow.getPhysicalNumberOfCells();
            for (int i = 0; i < columnCount; i++) {
                String cellName = headerRow.getCell(i).getStringCellValue();
                if (StringUtil.isBlank(cellName)) {
                    break;
                }
                factColumnCount++;
            }

            for (int i = 0; i < factColumnCount; i++) {
                String cellName = headerRow.getCell(i).getStringCellValue();
                if (!StringUtil.isBlank(cellName)) {
                    sqlFile.write("\"" + cellName + "\"");
                    if (i < factColumnCount - 1) {
                        sqlFile.write(", ");
                    }
                }
            }
            sqlFile.write(") VALUES \n");

            for (int r = 1; r < rowCount; r++) {
                Row row = sheet.getRow(r);
                sqlFile.write("(gen_random_uuid(), ");

                for (int c = 0; c < factColumnCount; c++) {
                    Cell cell = row.getCell(c);
                    if (cell == null || cell.getCellType() == CellType.BLANK) {
                        continue;
                    } else {
                        switch (cell.getCellType()) {
                            case STRING:
                                sqlFile.write("'" + cell.getStringCellValue() + "'");
                                break;
                            case NUMERIC:
                                if (DateUtil.isCellDateFormatted(cell)) {
                                    LocalDateTime localDateTimeCellValue = cell.getLocalDateTimeCellValue();
                                    sqlFile.write("'" + localDateTimeCellValue.toString() + "'");
                                } else {
                                    DataFormatter dataformatter = new DataFormatter();
                                    String cellValue = dataformatter.formatCellValue(cell);
                                    sqlFile.write(cellValue);

                                }
                                break;
                            case BOOLEAN:
                                sqlFile.write(String.valueOf(cell.getBooleanCellValue()));
                                break;
                            case FORMULA:
                                sqlFile.write("'" + cell.getCellFormula() + "'");
                                break;
                            default:

                        }
                    }
                    if (c < factColumnCount - 1) {
                        sqlFile.write(", ");
                    }
                }

                sqlFile.write(")");
                if (r < rowCount - 1) {
                    sqlFile.write(",\n");
                } else {
                    sqlFile.write(";");
                }
            }

            log.info("SQL file generated successfully!");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}