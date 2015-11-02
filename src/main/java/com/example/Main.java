package com.example;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.logging.Level;
import java.util.logging.Logger;

public class Main {
    static Logger logger = Logger.getLogger(Main.class.getName());

    public static void main(String[] args) {
        logger.setLevel(Level.ALL);
        logger.log(Level.INFO, "go");

        try (
                InputStream file = new FileInputStream(args[0]);
                Workbook wb = WorkbookFactory.create(file)
        ) {
            logger.log(Level.FINE, "go");
            for (Sheet sheet: wb) {
                logger.log(Level.FINE, sheet.getSheetName());
                for (Row row: sheet) {
                    logger.log(Level.FINER, "next row");
                    for (Cell cell: row) {
                        switch (cell.getCellType()) {
                            case Cell.CELL_TYPE_STRING:
                                logger.log(Level.FINER, cell.getStringCellValue());
                                break;
                            case Cell.CELL_TYPE_NUMERIC:
                                if (DateUtil.isCellDateFormatted(cell)) {
                                    logger.log(Level.FINER, cell.getDateCellValue().toString());
                                } else {
                                    logger.log(Level.FINER, "{0}", cell.getNumericCellValue());
                                }
                                break;
                        }
                    }
                }
            }
        } catch (FileNotFoundException ex) {
            logger.log(Level.SEVERE, "open file", ex);
        } catch (IOException|InvalidFormatException ex) {
            logger.log(Level.SEVERE, "read file", ex);
        }
    }
}
