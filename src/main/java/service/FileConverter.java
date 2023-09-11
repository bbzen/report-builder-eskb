package service;

import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.stream.Collectors;

public class FileConverter {
    private final DateTimeFormatter timeFormatter;
    private final DateTimeFormatter dateFormatter;

    public FileConverter() {
        this.timeFormatter = DateTimeFormatter.ofPattern("HH:mm");
        this.dateFormatter = DateTimeFormatter.ofPattern("dd.MM.yy");
    }

    @SneakyThrows
    public void convertToXls() {
        List<Path> foundFiles = findFile();
        if (foundFiles == null) {
            throw new RuntimeException("No files has been found");
        }
        for (Path file : foundFiles) {
            Document doc = Jsoup.parse(file.toFile());
            Element table = doc.getElementById("tableGraph2");
            if (table != null) {
            processTableElement(table, file);
            }
        }
    }

    @SneakyThrows
    private void processTableElement(Element table, Path file) {
        XSSFWorkbook book = new XSSFWorkbook();
        XSSFSheet sheet = book.createSheet("DataSheet");
        XSSFRow row = null;
        Cell cell;

        Elements headerOuter = table.getElementsByTag("thead");
        Elements headerOfTable = headerOuter.first().getElementsByTag("tr");
        Elements elementsOfHeader = headerOfTable.first().getElementsByTag("th");
        row = sheet.createRow(0);

        for (int i = 0; i < elementsOfHeader.size(); i++) {
            cell = row.createCell(i);
            Element element = elementsOfHeader.get(i);
            cell.setCellValue(element.text());
        }

        Elements tableBody = table.getElementsByTag("tbody");
        Elements tableBodyRows = tableBody.first().getElementsByTag("tr");
        for (int i = 0; i < tableBodyRows.size(); i++) {
            row = sheet.createRow(i + 1);
            Elements rowElements = tableBodyRows.get(i).getElementsByTag("td");
            for (int j = 0; j < rowElements.size(); j++) {
                cell = row.createCell(j);
                Element element = rowElements.get(j);
                if (j == 6) {
                    cell.setCellValue(String.valueOf(LocalDate.parse(element.text(), dateFormatter)));
                } else if (j < 5) {
                    cell.setCellValue(Double.parseDouble(element.text()));
                } else {
                    cell.setCellValue(element.text());
                }
            }
        }
        try (FileOutputStream fout = new FileOutputStream("./reports/" + file.getFileName() + ".xlsx")) {
            book.write(fout);
        }
    }

    private static List<Path> findFile() {
        List<Path> foundFiles = null;
        try {
            foundFiles = Files.list(Path.of("./reports"))
                    .filter(file -> !Files.isDirectory(file))
                    .filter(file -> !file.getFileName().endsWith("html"))
                    .collect(Collectors.toList());
        } catch (IOException e) {
            throw new RuntimeException("Something went wrong");
        }

        if (foundFiles != null) {
            for (Path file : foundFiles) {
                System.out.println(file.getFileName());
            }
        }
        return foundFiles;
    }
}
