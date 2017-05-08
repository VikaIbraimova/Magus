package ru.d1g;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import java.io.*;
import java.util.*;

/**
 * Created by A on 03.05.2017.
 */
@Component
public class Utils {

    @Value(value = "${import_folder}")
    private String importDirectory;
    @Value(value = "${output_file}")
    private String outputFile;
    @Value(value = "#{'${headers}'.split('\\s*,\\s*')}")
    private List<String> headers;

    private static Logger log = LoggerFactory.getLogger(Utils.class);

    /*
    *  Получаем мапу заголовок - номер столбца
    * */
    @Deprecated
    public Map<String, Integer> getHeadersColumnNumbers(Row row, String[] headers) {
        Map<String, Integer> map = new HashMap<>();

        row.cellIterator().forEachRemaining(cell -> {
            String cellValue = cell.getStringCellValue();

            for (String header : headers
                    ) {
                if (cellValue.equals(header)) {
                    int columnIndex = cell.getColumnIndex();
                    log.trace("found header in column {}", columnIndex);
                    map.put(header, columnIndex);
                }
            }
        });
        return map;
    }

    /*
    * Получить все возможные заголовки на строке
    * */
    public Map<String, Integer> getAllRowHeaders(Row row) {
        Map<String, Integer> map = new HashMap<>();

        row.cellIterator().forEachRemaining(cell -> {
            String cellStringValue = cell.getStringCellValue();
            if (!cellStringValue.equals("")) {
                map.put(cellStringValue, cell.getColumnIndex());
            }
        });
        return map;
    }

    /*
    *  Ищем начальную строку для обработки
    * */
    public Integer getStartingRow(Sheet sheet, String tag) {
        Integer startingRow = 0;
        for (Iterator<Row> it = sheet.rowIterator(); it.hasNext(); ) {
            Row row = it.next();
            Cell zeroCell = row.getCell(0);
            if (zeroCell != null) {
                if (zeroCell.getStringCellValue().equals(tag)) {
                    startingRow = row.getRowNum();
                    break;
                } else {
//                    throw new RuntimeException("can't find starting row");
                }
            }
        }
        return startingRow;
    }

    public List<String> getImportedFilesList() {
        List<String> files = new ArrayList<String>();

        File importDir = new File(importDirectory);
        String[] filenames = importDir.list();
        System.out.println(importDir.getAbsolutePath());
        for (String filename : filenames) {
            files.add(importDirectory + File.separator + filename);
        }

        return files;
    }

    public String getOutputFile() {
        File importDir = new File(outputFile);
        return importDir.getPath();
    }

    public XSSFWorkbook getWorkBookFromFile(String filepath) {
        try (InputStream is = new FileInputStream(filepath)) {
//                берем книгу
            return new XSSFWorkbook(is);
        } catch (FileNotFoundException e) {
            log.error("file not found: {}", filepath);
        } catch (IOException e) {
            log.error("IOException for: {}", filepath);
        }
        return null;
    }

    public void saveWorkbook(Workbook workbook, String file) {
        try (FileOutputStream outputStream = new FileOutputStream(file)) {
            workbook.write(outputStream);
        } catch (IOException e) {
            log.error("не удаётся сохранить книгу {}", file);
        }
    }

    public List<String> getHeaders() {
        return headers;
    }

    public boolean checkMapContainsEqualHeaders(Map<String, ?> inputFileHeadersMap, Map<String, ?> outputFileHeadersMap) {
        return outputFileHeadersMap.keySet().containsAll(inputFileHeadersMap.keySet());
    }

    /*
    * сравнение ячеек
    * */
    public boolean compareCells(Cell a, Cell b) {
        if (a != null && b != null) {
            if (a.getCellTypeEnum() == b.getCellTypeEnum()) {
                return (isCellEqual(a, b, a.getCellTypeEnum()));
            }
        }
        throw new RuntimeException("cell is null");
    }

    /*
    * проверка равности ячеек, требуется указать тип ячеек.
    * */
    private boolean isCellEqual(Cell a, Cell b, CellType cellType) {
        boolean result = false;
        switch (cellType) {
            case NUMERIC:
                result = a.getNumericCellValue() == b.getNumericCellValue();
                break;
            case BOOLEAN:
                result = a.getBooleanCellValue() == b.getBooleanCellValue();
                break;
            case STRING:
                result = a.getRichStringCellValue().getString().equals(b.getRichStringCellValue().getString());
                break;
            case FORMULA:
                result = a.getCellFormula().equals(b.getCellFormula());
                break;
            case ERROR:
                result = a.getErrorCellValue() == b.getErrorCellValue();
                break;
            case BLANK:
                result = a.getStringCellValue().equals(b.getStringCellValue());
                break;
        }
        return result;
    }

    /*
    * копирование ячеек
    * */
    public void copyCell(Cell fromCell, Cell toCell) {
        if (fromCell != null && toCell != null) {
            switch (fromCell.getCellTypeEnum()) {
                case NUMERIC:
                    toCell.setCellValue(fromCell.getNumericCellValue());
                    break;
                case FORMULA:
                    toCell.setCellFormula(fromCell.getCellFormula());
                    break;
                case STRING:
                    toCell.setCellValue(fromCell.getStringCellValue());
                    break;
                case BOOLEAN:
                    toCell.setCellValue(fromCell.getBooleanCellValue());
                    break;
                case ERROR:
                    toCell.setCellValue(fromCell.getErrorCellValue());
                    break;
                case BLANK:
                    toCell.setCellValue("");
                    break;
            }
        }
    }
}
