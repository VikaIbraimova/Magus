package ru.d1g.impl;

import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;
import ru.d1g.Utils;

import java.time.Duration;
import java.time.Instant;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;

/**
 * Created by A on 09.05.2017.
 */
@Component
public class Parser {

    private static Logger log = LoggerFactory.getLogger(Parser.class);
    private ExecutorService executorService = Executors.newSingleThreadExecutor();
    private Utils utils;
    private final ParserThreadFactory parserThreadFactory;

    private Workbook outputWorkbook;
    private Sheet outputFileSheet;
    private Map<String, Integer> outputFileHeadersMap;
    private Integer outputStartingRow;

    @Autowired
    public Parser(Utils utils, ParserThreadFactory parserThreadFactory) {
        this.utils = utils;
        this.parserThreadFactory = parserThreadFactory;
    }

    public void parse() throws Exception {
        List<String> importFiles = utils.getImportedFilesList();
        String outputFile = utils.getOutputFile();

        outputWorkbook = utils.getWorkBookFromFile(outputFile); // output книга
        outputFileSheet = outputWorkbook.getSheetAt(0); //лист с которым будем сравнивать все входящие файлы
        outputFileHeadersMap = utils.getAllRowHeaders(outputFileSheet.getRow(0)); // получаем номера столбцов со всеми заголовками у OUTPUT файла
        outputStartingRow = utils.getStartingRow(outputFileSheet, "start"); // номер строки с которой начинаем сравнение (помечен словом start)
        Instant start = Instant.now();
        for (String filePathString : importFiles
                ) {
            executorService.submit(parserThreadFactory.getObject().setFilePathString(filePathString));
        }
        executorService.shutdown();
        executorService.awaitTermination(1, TimeUnit.HOURS);
        Instant end = Instant.now();
        log.trace("execution time is: {}",Duration.between(start,end));
    }

    public Workbook getOutputWorkbook() {
        return outputWorkbook;
    }

    public Sheet getOutputFileSheet() {
        return outputFileSheet;
    }

    public Map<String, Integer> getOutputFileHeadersMap() {
        return outputFileHeadersMap;
    }

    public Integer getOutputStartingRow() {
        return outputStartingRow;
    }
}
