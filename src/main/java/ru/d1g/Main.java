package ru.d1g;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.context.ApplicationContext;
import org.springframework.context.support.ClassPathXmlApplicationContext;

import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * Created by A on 03.05.2017.
 */
public class Main {

    private static Logger log = LoggerFactory.getLogger(Main.class);
    private static ApplicationContext applicationContext = new ClassPathXmlApplicationContext("spring-config.xml");
    private static List<String> headers;

    public static void main(String[] args) {

        Utils utils = applicationContext.getBean(Utils.class);
        XSSFCellStyle coloredCellStyle;
        headers = utils.getHeaders();

        List<String> importFiles = utils.getImportedFilesList();
        String outputFile = utils.getOutputFile();

        XSSFWorkbook outputWorkbook = utils.getWorkBookFromFile(outputFile); // output книга
        XSSFSheet outputFileSheet = outputWorkbook.getSheetAt(0); //лист с которым будем сравнивать все входящие файлы
        Map<String, Integer> outputFileHeadersMap = utils.getAllRowHeaders(outputFileSheet.getRow(0)); // получаем номера столбцов со всеми заголовками у OUTPUT файла
        Integer outputStartingRow = utils.getStartingRow(outputFileSheet, "start"); // номер строки с которой начинаем сравнение (помечен словом start)


        for (String filePathString : importFiles
                ) {

            XSSFWorkbook importWorkbook = utils.getWorkBookFromFile(filePathString); // получаем книгу
            XSSFSheet importFileSheet = importWorkbook.getSheetAt(0); // берем лист
            XSSFRow headersRow = importFileSheet.getRow(0); // берем первую строку (должна быть с заголовками)
            Map<String, Integer> inputFileHeadersMap = utils.getAllRowHeaders(headersRow); // получаем номера стобцов со всеми заголовками входного файла
            Integer importStartingRow = utils.getStartingRow(importFileSheet, "start"); // номер строки с которой начинаем сравнение (помечен словом start)
            coloredCellStyle = importWorkbook.createCellStyle();
            coloredCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            coloredCellStyle.setFillForegroundColor(IndexedColors.CORAL.getIndex());

            if (!utils.checkMapContainsEqualHeaders(inputFileHeadersMap, outputFileHeadersMap)) { // проверим, что все заголовки входного файла так-же присутствуют и в выходном файле
                log.error("в выходном файле отсутствует заголовок из входного файла {}",filePathString);
                throw new RuntimeException("в выходном файле отсутствует заголовок из входного файла");
            }
            boolean rowFound = false; // состояние, по которому будем определять, что строка не была найдена

            for (int importRowNum = importStartingRow; importRowNum < importFileSheet.getPhysicalNumberOfRows(); importRowNum++) { // для каждой строки import файла
                XSSFRow importRow = importFileSheet.getRow(importRowNum);
                rowFound = false; // обнуляем состояние
                for (int outputRowNum = outputStartingRow; outputRowNum < outputFileSheet.getPhysicalNumberOfRows(); outputRowNum++) { // для каждой строки out файла
                    XSSFRow outputRow = outputFileSheet.getRow(outputRowNum);

                    int equalCounter = 0;  // заводим счётчик совпавших заголовков
                    for (String header : headers // сравниваем значения нужных заголовков
                            ) {

                        XSSFCell outputCell = outputRow.getCell(outputFileHeadersMap.get(header));
                        XSSFCell importCell = importRow.getCell(inputFileHeadersMap.get(header));

                        if (utils.compareCells(outputCell, importCell)) {
                            equalCounter++; // нашли совпадение по колонке, следовательно инкрементируем счетчик
                        }
                    }

                    if (equalCounter == headers.size()) { // если счетчик совпадений равен кол-ву сравниваемых заголовков -> значит мы нашли совпадение строк
                        log.trace("найдено совпадение строк в файлах import: {} output: {}. заполняем данными", importRow.getRowNum(), outputRow.getRowNum());
                        inputFileHeadersMap.forEach((header, columnNumber) -> { // вот такую жуть ещё делаем :D для каждого заголовка:
                                XSSFCell outputCell = outputRow.getCell(outputFileHeadersMap.get(header)); // берем ячейку выходного файла для заголовка
                                if (outputCell == null) { // проверяем, если в выходном файле ячейка отсутствует, то создаём её
                                    outputRow.createCell(outputFileHeadersMap.get(header));
                                    outputCell = outputRow.getCell(outputFileHeadersMap.get(header));
                                }
                                utils.copyCell(importRow.getCell(columnNumber),outputCell);
                        });
                        rowFound = true; // меняем состояние: найдено соответствие строк
                    }
                }

                if (!rowFound) {
                    log.trace("строка {} файла {} не была найдена, красим", importRow.getRowNum(), filePathString);
                    for (Iterator<Cell> it = importRow.cellIterator(); it.hasNext(); ) {
                        Cell cell = it.next();
                        cell.setCellStyle(coloredCellStyle);
                    }
                }
            }
            log.trace("сохраняем import книгу");
            utils.saveWorkbook(importWorkbook, filePathString);
        }
        log.trace("сохраняем output книгу");
        utils.saveWorkbook(outputWorkbook, outputFile);
    }
}