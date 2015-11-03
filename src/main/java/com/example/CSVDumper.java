package com.example;

import com.opencsv.CSVWriter;

import java.io.Closeable;
import java.io.IOException;
import java.io.Writer;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

public class CSVDumper implements Closeable, Dumper {
  private final String imagesDir;
  private ArrayList<String> currentRow;
  private CSVWriter out;

  public CSVDumper(Writer os, Path imagesDir) {
    this.out = new CSVWriter(os);
    this.imagesDir = imagesDir.toString();
    this.currentRow = new ArrayList<>(0);
  }

  @Override
  public void close() throws IOException {
    if (out != null) {
      flushLine();
      out.close();
    }
  }

  @Override
  public void writePictures(List<String> pictures) {
    String val = pictures.stream()
        .map(pic -> Paths.get(imagesDir, pic).toString())
        .collect(Collectors.joining(";"));
    writeCell(val);
  }

  @Override
  public void writeCell(String val) {
    currentRow.add(val);
  }

  @Override
  public void flushLine() {
    if (currentRow.size() == 0) {
      return;
    }

    String[] stringEntries = currentRow.toArray(new String[currentRow.size()]);
    out.writeNext(stringEntries);
    currentRow = new ArrayList<>(0);
  }
}
