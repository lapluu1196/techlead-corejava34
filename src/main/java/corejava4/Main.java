package corejava4;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public class Main {
    public static final int COLUMN_INDEX_ID = 1;
    public static final int COLUMN_INDEX_NAME = 2;

    public static void main(String[] args) {
        final String excelPath = System.getProperty("user.dir") + "\\src\\main\\resources\\excel\\BangCong.xlsx";

        try (InputStream is = new FileInputStream(new File(excelPath))) {
            List<Employee> employees = new ArrayList<>();

            Workbook workbook = getWorkbook(is, excelPath);

            Sheet sheet = workbook.getSheetAt(0);

            // Identify "Tổng lương" column
            int totalSalaryColumn = 0;
            for (int i = 0; i < sheet.getRow(3).getLastCellNum(); i++) {
                Cell cell = sheet.getRow(3).getCell(i);
                if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().equalsIgnoreCase("Tổng lương")) {
                    totalSalaryColumn = cell.getColumnIndex();
                }
            }

            // Get Map of Day Range
            Map<Integer, List<Integer>> colDayRange = new LinkedHashMap<>();
            Row dayRow = sheet.getRow(3);
            int currentDay = -1;
            List<Integer> dayRange = new ArrayList<>();

            for (int i = totalSalaryColumn + 1; i < dayRow.getLastCellNum(); i++) {
                Cell cell = dayRow.getCell(i);
                if (cell != null) {
                    String cellValue = "";
                    if (cell.getCellType() == CellType.NUMERIC) {
                        cellValue = String.valueOf((int) cell.getNumericCellValue());
                    } else if (cell.getCellType() == CellType.STRING) {
                        cellValue = cell.getStringCellValue().trim();
                    }

                    try {
                        int dayNumber = Integer.parseInt(cellValue);
                        if (dayNumber != currentDay) {
                            if (currentDay != -1 && !dayRange.isEmpty()) {
                                colDayRange.put(currentDay, new ArrayList<>(dayRange));
                                dayRange.clear();
                            }
                            currentDay = dayNumber;
                        }
                        dayRange.add(i);
                    } catch (NumberFormatException e) {
                        if (currentDay != -1) {
                            dayRange.add(i);
                        }
                    }
                }
            }
            if (!dayRange.isEmpty()) {
                colDayRange.put(currentDay, dayRange);
            }
            colDayRange.forEach((k, v) -> System.out.println(k + ": " + v));

            //Get Employee data
            for (int i = 6; i < sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null || isRowEmpty(row)) continue;
                String employeeId = row.getCell(COLUMN_INDEX_ID).getStringCellValue();
                String employeeName = row.getCell(COLUMN_INDEX_NAME).getStringCellValue();

                // Get shifts rate
                Map<String, Double> shiftsRate = new LinkedHashMap<>();
                for (int j = 3; j < totalSalaryColumn; j++) {
                    String shiftName = sheet.getRow(5).getCell(j).getStringCellValue().trim();
                    if (shiftName != null && !shiftName.equalsIgnoreCase("$")) {
                        if (shiftName.contains("WK")) {
                            shiftsRate.put(shiftName, sheet.getRow(i).getCell(totalSalaryColumn - 1).getNumericCellValue());
                        } else {
                            shiftsRate.put(shiftName, sheet.getRow(i).getCell(j + 1).getNumericCellValue());
                        }
                    }
                }

                // Get working days
                List<WorkingDay> workingDays = new ArrayList<>();
                for (Map.Entry<Integer, List<Integer>> entry : colDayRange.entrySet()) {
                    WorkingDay workingDay = new WorkingDay(String.valueOf(entry.getKey()));

                    for (int j = 0; j < entry.getValue().size(); j++) {
                        int colIndex = entry.getValue().get(j);
                        String shiftName = sheet.getRow(5).getCell(colIndex).getStringCellValue();

                        if (shiftName != null && !shiftName.equalsIgnoreCase("$")) {
                            double hours = 0.0;
                            Cell shiftCell = row.getCell(colIndex);
                            if (shiftCell != null && shiftCell.getCellType() == CellType.NUMERIC) {
                                hours = shiftCell.getNumericCellValue();
                            }
                            if (!shiftsRate.containsKey(shiftName)) {
                                System.err.println("The shift '" + shiftName + "' has no rate for employees: " + employeeName);
                                continue;
                            }
                            double rate = shiftsRate.getOrDefault(shiftName, 0.0);
                            double amount = rate * hours;
                            Shift shift = new Shift(shiftName, hours, amount);
                            workingDay.addShift(shift, hours, amount);
                        }
                    }
                    workingDays.add(workingDay);
                }

                Employee employee = new Employee(employeeId, employeeName, sheet.getRow(i).getCell(totalSalaryColumn).getNumericCellValue());
                employee.setShiftsRate(shiftsRate);
                for (WorkingDay workingDay : workingDays) {
                    employee.addWorkingDay(workingDay);
                }

                employees.add(employee);
            }
            for (Employee employee : employees) {
                System.out.println(employee.toString());
            }

        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private static Workbook getWorkbook(InputStream is, String excelPath) throws IOException {
        Workbook workbook = null;
        if (excelPath.endsWith("xlsx")) {
            workbook = new XSSFWorkbook(is);
        } else if (excelPath.endsWith("xls")) {
            workbook = new HSSFWorkbook(is);
        } else {
            throw new IllegalArgumentException("The specified file is not Excel file.");
        }
        return workbook;
    }

    private static Object getCellValue(Cell cell) {
        CellType cellType = cell.getCellType();
        Object cellValue = null;
        switch (cellType) {
            case STRING:
                cellValue = cell.getStringCellValue();
                break;
            case NUMERIC:
                cellValue = cell.getNumericCellValue();
                break;
            case BOOLEAN:
                cellValue = cell.getBooleanCellValue();
                break;
            case FORMULA:
                Workbook workbook = cell.getSheet().getWorkbook();
                FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                cellValue = evaluator.evaluate(cell).getNumberValue();
                break;
            case _NONE:
            case BLANK:
            case ERROR:
            default:
                return null;
        }

        return cellValue;
    }

    private static boolean isRowEmpty(Row row) {
        if (row == null) return true;
        for (int cellNum = 0; cellNum < row.getLastCellNum(); cellNum++) {
            Cell cell = row.getCell(cellNum);
            if (cell != null && cell.getCellType() != CellType.BLANK) {
                return false;
            }
        }
        return true;
    }
}
