package com.example;

import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Path;

public class SaverFactory {
  public static Saver create(Path path, Path imagesDir) throws IOException {
    if (path.toString().endsWith(".html")) {
      return new HTMLSaver(new FileWriter(path.toString()), imagesDir);
    } else {
      return new CSVSaver(new FileWriter(path.toString()), imagesDir);
    }
  }
}
