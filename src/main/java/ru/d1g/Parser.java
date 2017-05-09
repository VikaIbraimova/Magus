package ru.d1g;

import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;
import ru.d1g.exceptions.ParseException;

import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * Created by A on 09.05.2017.
 */
@Component
public class Parser {

    private static Logger log = LoggerFactory.getLogger(Parser.class);

    private Utils utils;

    @Autowired
    public Parser(Utils utils) {
        this.utils = utils;
    }

    public void parse() {
        CellStyle coloredCellStyle;
        List<String> headers = utils.getHeaders();

        List<String> importFiles = utils.getImportedFilesList();
        String outputFile = utils.getOutputFile();

        Workbook outputWorkbook = utils.getWorkBookFromFile(outputFile); // output книга
        Sheet outputFileSheet = outputWorkbook.getSheetAt(0); //лист с которым будем сравнивать все входящие файлы
        Map<String, Integer> outputFileHeadersMap = utils.getAllRowHeaders(outputFileSheet.getRow(0)); // получаем номера столбцов со всеми заголовками у OUTPUT файла
        Integer outputStartingRow = utils.getStartingRow(outputFileSheet, "start"); // номер строки с которой начинаем сравнение (помечен словом start)


        for (String filePathString : importFiles
                ) {
            Workbook importWorkbook = utils.getWorkBookFromFile(filePathString); // получаем книгу

            Sheet importFileSheet;
            for (Iterator<Sheet> sheetIterator = importWorkbook.sheetIterator(); sheetIterator.hasNext(); ) { // для каждого листа
                importFileSheet = sheetIterator.next();

                Row headersRow = importFileSheet.getRow(0); // берем первую строку (должна быть с заголовками)
                Map<String, Integer> inputFileHeadersMap = utils.getAllRowHeaders(headersRow); // получаем номера стобцов со всеми заголовками входного файла
                Integer importStartingRow = utils.getStartingRow(importFileSheet, "start"); // номер строки с которой начинаем сравнение (помечен словом start)
                coloredCellStyle = importWorkbook.createCellStyle();
                coloredCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                coloredCellStyle.setFillForegroundColor(IndexedColors.CORAL.getIndex());

                if (!utils.checkMapContainsEqualHeaders(inputFileHeadersMap, outputFileHeadersMap)) { // проверим, что все заголовки входного файла так-же присутствуют и в выходном файле
                    log.error("в выходном файле отсутствует заголовок из входного файла {}", filePathString);
                    throw new ParseException("в выходном файле отсутствует заголовок \""
                            + Arrays.toString(inputFileHeadersMap.keySet().stream().filter(s -> !outputFileHeadersMap.containsKey(s)).toArray())
                            + "\"\nиз входного файла:\n"
                            + filePathString
                            + "\nна листе:\n"
                            + importFileSheet.getSheetName());
                }
                boolean rowFound = false; // состояние, по которому будем определять, что строка не была найдена

                for (int importRowNum = importStartingRow; importRowNum < importFileSheet.getPhysicalNumberOfRows(); importRowNum++) { // для каждой строки import файла
                    Row importRow = importFileSheet.getRow(importRowNum);
                    if (importRow==null) continue;
                    rowFound = false; // обнуляем состояние
                    for (int outputRowNum = outputStartingRow; outputRowNum < outputFileSheet.getPhysicalNumberOfRows(); outputRowNum++) { // для каждой строки out файла
                        Row outputRow = outputFileSheet.getRow(outputRowNum);
                        if (outputRow==null) continue;

                        int equalCounter = 0;  // заводим счётчик совпавших заголовков
                        for (String header : headers // сравниваем значения нужных заголовков
                                ) {

                            Cell outputCell = outputRow.getCell(outputFileHeadersMap.get(header));
                            Cell importCell = importRow.getCell(inputFileHeadersMap.get(header));

                            if (utils.compareCells(outputCell, importCell)) {
                                equalCounter++; // нашли совпадение по колонке, следовательно инкрементируем счетчик
                            }
                        }

                        if (equalCounter == headers.size()) { // если счетчик совпадений равен кол-ву сравниваемых заголовков -> значит мы нашли совпадение строк
                            log.trace("найдено совпадение строк в файлах import: {} output: {}. заполняем данными", importRow.getRowNum(), outputRow.getRowNum());
                            inputFileHeadersMap.forEach((header, columnNumber) -> { // вот такую жуть ещё делаем :D для каждого заголовка:
                                Cell outputCell = outputRow.getCell(outputFileHeadersMap.get(header)); // берем ячейку выходного файла для заголовка
                                if (outputCell == null) { // проверяем, если в выходном файле ячейка отсутствует, то создаём её
                                    outputRow.createCell(outputFileHeadersMap.get(header));
                                    outputCell = outputRow.getCell(outputFileHeadersMap.get(header));
                                }
                                utils.copyCell(importRow.getCell(columnNumber), outputCell);
                            });
                            rowFound = true; // меняем состояние: найдено соответствие строк
                        }
                    }

                    if (!rowFound) {
                        log.trace("строка {} файла {} на листе {} не была найдена, красим", importRow.getRowNum(), filePathString, importFileSheet.getSheetName());
                        for (Iterator<Cell> cellIterator = importRow.cellIterator(); cellIterator.hasNext(); ) {
                            Cell cell = cellIterator.next();
                            cell.setCellStyle(coloredCellStyle);
                        }
                    }
                }
                utils.saveWorkbook(importWorkbook, filePathString);
                utils.saveWorkbook(outputWorkbook, outputFile);
            }
        }
    }
}
