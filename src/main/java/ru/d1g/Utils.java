package ru.d1g;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;
import ru.d1g.exceptions.ParseException;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
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
            if (cell != null && cell.getCellTypeEnum() == CellType.STRING) {
                String cellStringValue = cell.getStringCellValue();
                if (!cellStringValue.equals("")) {
                    map.put(cellStringValue, cell.getColumnIndex());
                }
            }
        });
        return map;
    }

    /*
    *  Ищем начальную строку для обработки
    * */
    public Integer getStartingRow(Sheet sheet, String tag) {
        Integer startingRow = -1;
        for (Iterator<Row> it = sheet.rowIterator(); it.hasNext(); ) {
            Row row = it.next();
            Cell zeroCell = row.getCell(0);
            if (zeroCell != null && zeroCell.getCellTypeEnum() == CellType.STRING) {
                if (zeroCell.getStringCellValue().equals(tag)) {
                    startingRow = row.getRowNum();
                    break;
                }
            }
        }
        return startingRow;
    }

    public List<String> getImportedFilesList() {
        List<String> filesStringsList = new ArrayList<>();
        try {
            Files.newDirectoryStream(Paths.get(importDirectory)).forEach(path -> {
                if (path.toFile().isFile()) {
                    filesStringsList.add(path.toString());
                }
            });
        } catch (IOException e) {
            log.error("не удается открыть файл", e);
            throw new ParseException("can't open file\n", e);
        }
        return filesStringsList;
    }

    public String getOutputFile() {
        File file = Paths.get(outputFile).toFile();
        return file.toString();
    }

    public XSSFWorkbook getWorkBookFromFile(String filepath) {
        try (InputStream is = new FileInputStream(filepath)) {
//                берем книгу
            return new XSSFWorkbook(is);
        } catch (FileNotFoundException e) {
            log.error("file not found: {}", filepath);
            throw new ParseException("file not found: " + filepath + "\n", e);
        } catch (IOException e) {
            log.error("IOException for: {}", filepath);
            throw new ParseException("unexpected IO error while reading file " + filepath + "\n", e);
        }
    }

    public void saveWorkbook(Workbook workbook, String file) {
        try (FileOutputStream outputStream = new FileOutputStream(file)) {
            log.trace("сохраняем файл: {}", file);
            workbook.write(outputStream);
        } catch (IOException e) {
            log.error("не удаётся сохранить книгу {}", file);
            throw new ParseException("can't save book\n", e);
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
                return (isCellEqual(a, b, a.getCellTypeEnum(), false, null));
            }
        }
        return false;
    }

    public boolean compareCells(Cell a, Cell b, boolean ignoreCase, String regex) {
        if (a != null && b != null) {
            if (a.getCellTypeEnum() == b.getCellTypeEnum()) {
                return (isCellEqual(a, b, a.getCellTypeEnum(), ignoreCase, regex));
            }
        }
        return false;
    }

    /*
    * проверка равности ячеек, требуется указать тип ячеек.
    * */
    private boolean isCellEqual(Cell a, Cell b, CellType cellType, boolean ignoreCase, String regex) {
        boolean result = false;
        switch (cellType) {
            case NUMERIC:
                result = a.getNumericCellValue() == b.getNumericCellValue();
                break;
            case BOOLEAN:
                result = a.getBooleanCellValue() == b.getBooleanCellValue();
                break;
            case STRING:
                String x = a.getRichStringCellValue().getString();
                String y = b.getRichStringCellValue().getString();
                if (regex != null) {
                    x = x.replaceAll(regex, "");
                    y = y.replaceAll(regex, "");
                }
                if (ignoreCase) {
                    result = x.equalsIgnoreCase(y);
                } else {
                    result = x.equals(y);
                }
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
