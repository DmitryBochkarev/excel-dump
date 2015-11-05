package com.example;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.stream.Collectors;

public class Main {
  private static final Logger logger = Logger.getLogger(Main.class.getName());

  public static void main(String[] args) throws IOException {
    logger.setLevel(Level.ALL);

    Path input;
    Path output;

    try {
      input = Paths.get(args[0]);
      output = Paths.get(args[1]);
    } catch (ArrayIndexOutOfBoundsException ex) {
      System.err.println("usage: excel-export input output");
      return;
    }

    Path outputDir = output.getParent();
    Path imagesDir = Paths.get(output.toString() + "_images");

    Files.createDirectories(outputDir);
    Files.createDirectories(imagesDir);

    try (
        InputStream file = new FileInputStream(input.toString());
        Workbook wb = WorkbookFactory.create(file);
        Saver out = SaverFactory.create(output, imagesDir)
    ) {
      for (Sheet sheet : wb) {
        dumpPictures(imagesDir, sheet);

        int rowStart = sheet.getFirstRowNum();
        int rowEnd = sheet.getLastRowNum();

        for (int rowNum = rowStart; rowNum <= rowEnd; rowNum++) {
          logger.log(Level.INFO, "row {0}", rowNum);
          Row row = sheet.getRow(rowNum);

          if (row == null) {
            continue;
          }

          int cellStart = row.getFirstCellNum();
          int cellEnd = row.getLastCellNum();

          for (int cellNum = cellStart; cellNum < cellEnd; cellNum++) {
            Cell cell = row.getCell(cellNum, Row.CREATE_NULL_AS_BLANK);

            logger.log(Level.INFO, "cell [{0}, {1}] format: {2}", new Object[]{rowNum, cellNum, cell.getCellStyle().getDataFormatString()});
            switch (cell.getCellType()) {
              case Cell.CELL_TYPE_STRING: {
                String val = cell.getStringCellValue();

                if (val.equals("")) {
                  List<String> pictures = getPictures(wb, sheet, cell);
                  out.writePictures(pictures);
                } else {
                  out.writeCell(val);
                }

                break;
              }
              case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                  out.writeCell(cell.getDateCellValue().toString());
                } else {
                  out.writeCell(String.valueOf(cell.getNumericCellValue()));
                }
                break;
              case Cell.CELL_TYPE_BOOLEAN:
                out.writeCell(String.valueOf(cell.getBooleanCellValue()));
                break;
              case Cell.CELL_TYPE_FORMULA:
                out.writeCell(cell.getCellFormula());
              default: {
                List<String> pictures = getPictures(wb, sheet, cell);
                out.writePictures(pictures);
              }
            }
          }

          out.flushLine();
        }
      }
    } catch (FileNotFoundException ex) {
      logger.log(Level.SEVERE, "open file", ex);
    } catch (IOException | InvalidFormatException ex) {
      logger.log(Level.SEVERE, "read file", ex);
    }
  }

  private static List<String> getPictures(Workbook wb, Sheet sheet, Cell cell) {
    if (wb instanceof HSSFWorkbook) {
      return getHSSFPictures((HSSFSheet) sheet, (HSSFCell) cell);
    } else {
      return getXSSFPictures((XSSFSheet) sheet, (XSSFCell) cell);
    }
  }

  private static List<String> getHSSFPictures(HSSFSheet sheet, HSSFCell cell) {
    List<String> pictures = new ArrayList<>(0);

    HSSFPatriarch drawing = sheet.getDrawingPatriarch();
    List<HSSFShape> shapes = drawing.getChildren();
    for (HSSFShape shape : shapes) {
      if (!(shape instanceof HSSFPicture)) {
        continue;
      }

      HSSFPicture picture = (HSSFPicture) shape;
      HSSFClientAnchor anchor = picture.getClientAnchor();
      if ((anchor.getCol1() <= cell.getColumnIndex()) &&
          (anchor.getRow1() <= cell.getRowIndex()) &&
          (anchor.getCol2() >= cell.getColumnIndex()) &&
          (anchor.getRow2() >= cell.getRowIndex())) {
        String filename = getHSSFPictureFilename(sheet, picture);
        pictures.add(filename);
      }
    }

    return pictures;
  }

  private static List<String> getXSSFPictures(XSSFSheet sheet, XSSFCell cell) {
    List<String> pictures = new ArrayList<>(0);

    XSSFDrawing drawing = sheet.getDrawingPatriarch();
    List<XSSFShape> shapes = drawing.getShapes();

    for (XSSFShape shape : shapes) {
      if (!(shape instanceof XSSFPicture)) {
        continue;
      }

      XSSFPicture picture = (XSSFPicture) shape;
      XSSFClientAnchor anchor = picture.getClientAnchor();

      if ((anchor.getCol1() <= cell.getColumnIndex()) &&
          (anchor.getRow1() <= cell.getRowIndex()) &&
          (anchor.getCol2() >= cell.getColumnIndex()) &&
          (anchor.getRow2() >= cell.getRowIndex())) {
        String filename = getXSSFPictureFilename(sheet, picture);
        pictures.add(filename);
      }
    }

    return pictures;
  }

  private static void dumpPictures(Path path, Sheet sheet) {
    if (sheet instanceof HSSFSheet) {
      dumpHSSFPictures(path.toString(), (HSSFSheet) sheet);
    } else {
      dumpXSSFPictures(path.toString(), (XSSFSheet) sheet);
    }
  }

  private static void dumpHSSFPictures(String path, HSSFSheet sheet) {
    HSSFPatriarch drawing = sheet.getDrawingPatriarch();
    List<HSSFShape> shapes = drawing.getChildren();
    for (HSSFShape shape : shapes) {
      if (!(shape instanceof HSSFPicture)) {
        continue;
      }

      HSSFPicture picture = (HSSFPicture) shape;
      String filename = getHSSFPictureFilename(sheet, picture);
      String finalpath = Paths.get(path, filename).toString();

      try (FileOutputStream out = new FileOutputStream(finalpath)) {
        out.write(picture.getPictureData().getData());
      } catch (IOException e) {
        logger.log(Level.SEVERE, "saving picture " + finalpath, e);
      }
    }
  }

  private static void dumpXSSFPictures(String path, XSSFSheet sheet) {
    XSSFDrawing drawing = sheet.getDrawingPatriarch();
    List<XSSFShape> shapes = drawing.getShapes();
    for (XSSFShape shape : shapes) {
      if (!(shape instanceof XSSFPicture)) {
        continue;
      }

      XSSFPicture picture = (XSSFPicture) shape;
      String filename = getXSSFPictureFilename(sheet, picture);
      String finalpath = Paths.get(path, filename).toString();

      try (FileOutputStream out = new FileOutputStream(finalpath)) {
        out.write(picture.getPictureData().getData());
      } catch (IOException e) {
        logger.log(Level.SEVERE, "saving picture " + finalpath, e);
      }
    }
  }

  private static String getHSSFPictureFilename(HSSFSheet sheet, HSSFPicture picture) {
    return sanitizeFileName(sheet.getSheetName() + "_" + picture.getPictureIndex() + "_" +
                                picture.getFileName() + "." + picture.getPictureData().suggestFileExtension());
  }

  private static String getXSSFPictureFilename(XSSFSheet sheet, XSSFPicture picture) {
    String blipId = picture.getCTPicture().getBlipFill().getBlip().getEmbed();
    return sanitizeFileName(sheet.getSheetName() + "_" + blipId + "." + picture.getPictureData().suggestFileExtension());
  }

  private static String sanitizeFileName(String name) {
    return name.chars()
               .mapToObj(i -> (char) i)
               .map(c -> Character.isWhitespace(c) ? '_' : c)
               .filter(c -> Character.isLetterOrDigit(c) || c == '-' || c == '_' || c == '.')
               .map(String::valueOf)
               .collect(Collectors.joining());
  }
}
