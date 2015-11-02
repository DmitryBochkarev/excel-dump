package com.example;

import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Path;

public class DumperFactory {
  public static Dumper create(Path path, Path imagesDir) throws IOException {
    if (path.toString().endsWith(".html")) {
      return new HTMLDumper(new FileWriter(path.toString()), imagesDir);
    } else {
      return new CSVDumper(new FileWriter(path.toString()), imagesDir);
    }
  }
}
