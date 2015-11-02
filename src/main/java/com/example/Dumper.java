package com.example;

import java.io.Closeable;
import java.util.List;

public interface Dumper extends Closeable {

  void writePictures(List<String> pictures);

  void writeCell(String val);

  void flushLine();
}
