package com.example;

import j2html.tags.Tag;

import java.io.Closeable;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.Writer;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

import static j2html.TagCreator.*;

public class HTMLSaver implements Closeable, Saver {
  private final String imagesDir;
  private ArrayList<Tag> currentRow;
  private ArrayList<Tag> rows;

  private PrintWriter out;

  public HTMLSaver(Writer os, Path imagesDir) {
    this.out = new PrintWriter(os);
    this.imagesDir = imagesDir.toString();
    this.rows = new ArrayList<>(0);
    this.currentRow = new ArrayList<>(0);
  }

  @Override
  public void close() throws IOException {
    if (out == null) {
      return;
    }

    flushLine();
    out.write("<!DOCTYPE html>");
    out.write(html().with(
        head().with(
            meta().withContent("text/html; charset=UTF-8")
                .attr("http-equiv", "content-type"),
            style().withText("table,tr,td {border: 1px solid grey;}")
        ),
        body().with(
            table().with(
                tbody().with(rows)
            )
        )
    ).render());
    out.close();
  }

  @Override
  public void writePictures(List<String> pictures) {
    currentRow.add(
            td().with(
                    pictures.stream()
                            .map(pic ->
                                    img().withSrc(Paths.get(imagesDir).getFileName().toString() + "/" + pic)
                                            .attr("style", "max-width: 300px; max-height: 300px"))
                            .collect(Collectors.toList())
            )
    );
  }

  @Override
  public void writeCell(String val) {
    currentRow.add(td(val));
  }

  @Override
  public void flushLine() {
    if (currentRow.size() > 0) {
      rows.add(tr().with(currentRow));
      currentRow = new ArrayList<>(0);
    }
  }
}
