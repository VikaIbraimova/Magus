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
        CellStyle coralCellStyle;
        CellStyle yellowCellStyle;
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
                Integer importStartingRow = utils.getStartingRow(importFileSheet, "start"); // номер строки с которой начинаем сравнение (помечен словом start)
                if (importStartingRow < 0) {
                    continue;
                }// если нет столбца со start -> следующая итерация
                Row headersRow = importFileSheet.getRow(0); // берем первую строку (должна быть с заголовками)
                Map<String, Integer> inputFileHeadersMap = utils.getAllRowHeaders(headersRow); // получаем номера стобцов со всеми заголовками входного файла

                coralCellStyle = importWorkbook.createCellStyle();
                coralCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                coralCellStyle.setFillForegroundColor(IndexedColors.CORAL.getIndex());

                yellowCellStyle = importWorkbook.createCellStyle();
                yellowCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                yellowCellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());

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
                boolean regexedRowFound = false;

                for (Iterator<Row> importRowIterator = importFileSheet.rowIterator(); importRowIterator.hasNext(); ) {
                    Row importRow = importRowIterator.next();
                    while (importRow.getRowNum() < importStartingRow) { // прогоняем итератор до start строки
                        importRow = importRowIterator.next();
                    }

                    rowFound = false; // обнуляем состояние
                    regexedRowFound = false;
                    for (Iterator<Row> outputRowIterator = outputFileSheet.rowIterator(); outputRowIterator.hasNext(); ) {
                        Row outputRow = outputRowIterator.next();
                        while (outputRow.getRowNum() < outputStartingRow) { // прогоняем итератор до start строки
                            outputRow = outputRowIterator.next();
                        }

                        int equalCounter = 0;  // заводим счётчик совпавших заголовков
                        int regexedEqualCounter = 0;
                        for (String header : headers // сравниваем значения нужных заголовков
                                ) {

                            Cell outputCell = outputRow.getCell(outputFileHeadersMap.get(header));
                            Cell importCell = importRow.getCell(inputFileHeadersMap.get(header));

                            if (utils.compareCells(outputCell, importCell)) {
                                equalCounter++; // нашли совпадение по колонке, следовательно инкрементируем счетчик
                            }
                            if (utils.compareCells(outputCell, importCell, true, "[\\-\\+\\.\\^:,\\s]")) {
                                regexedEqualCounter++; // нашли совпадение по колонке, следовательно инкрементируем счетчик
                            }
                        }

                        if (equalCounter == headers.size()) { // если счетчик совпадений равен кол-ву сравниваемых заголовков -> значит мы нашли совпадение строк
                            log.trace("найдено совпадение строк в файлах import: {} output: {}. заполняем данными", importRow.getRowNum(), outputRow.getRowNum());
                            copyHeadedCells(importRow, outputRow, inputFileHeadersMap, outputFileHeadersMap);
                            rowFound = true; // меняем состояние: найдено соответствие строк
                        }
                        if (regexedEqualCounter == headers.size()) {
                            log.trace("найдено regex совпадение строк в файлах import: {} output: {}. заполняем данными", importRow.getRowNum(), outputRow.getRowNum());
                            copyHeadedCells(importRow, outputRow, inputFileHeadersMap, outputFileHeadersMap);
                            regexedRowFound = true; // меняем состояние: найдено соответствие строк
                        }
                    }

                    if (!rowFound) {
                        log.trace("строка {} файла {} на листе {} не была найдена, красим", importRow.getRowNum(), filePathString, importFileSheet.getSheetName());
                        for (Iterator<Cell> cellIterator = importRow.cellIterator(); cellIterator.hasNext(); ) {
                            Cell cell = cellIterator.next();
                            cell.setCellStyle(yellowCellStyle);
                        }
                    }

                    if (!regexedRowFound) {
                        log.trace("regexed строка {} файла {} на листе {} не была найдена, красим", importRow.getRowNum(), filePathString, importFileSheet.getSheetName());
                        for (Iterator<Cell> cellIterator = importRow.cellIterator(); cellIterator.hasNext(); ) {
                            Cell cell = cellIterator.next();
                            cell.setCellStyle(coralCellStyle);
                        }
                    }
                }
                utils.saveWorkbook(importWorkbook, filePathString);
                utils.saveWorkbook(outputWorkbook, outputFile);
            }
        }
    }

    private void copyHeadedCells(Row importRow, Row outputRow, Map<String, Integer> inputFileHeadersMap, Map<String, Integer> outputFileHeadersMap) {
        inputFileHeadersMap.forEach((header, columnNumber) -> { // вот такую жуть ещё делаем :D для каждого заголовка:
            Cell outputCell = outputRow.getCell(outputFileHeadersMap.get(header)); // берем ячейку выходного файла для заголовка
            if (outputCell == null) { // проверяем, если в выходном файле ячейка отсутствует, то создаём её
                outputRow.createCell(outputFileHeadersMap.get(header));
                outputCell = outputRow.getCell(outputFileHeadersMap.get(header));
            }
            utils.copyCell(importRow.getCell(columnNumber), outputCell);
        });
    }
}
