package org.example;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartPanel;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.data.category.DefaultCategoryDataset;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;


import javax.swing.*;
import java.io.*;
import java.util.*;

/**
 * The main class of the application is for analyzing and creating a schedule of student progress.
 */

public class App {

    /** Assessment for an excellent student */
    private static final int EXCELLENT_GRADE = 5;
    /** Assessment for the good guy */
    private static final int GOOD_GRADE = 4;
    /** Assessment for a triple */
    private static final int SATISFACTORY_GRADE = 3;
    /**Logger instance */
    private static final Logger logger = LogManager.getLogger(App.class);

    /**
     * The entry point to the application.
     *
     * @param args command line arguments
     */

    public static void main(String[] args) {
        logger.info("The application is being launched");
        String inputFilePath = "C:\\Users\\Professional\\IdeaProjects\\Kursovaya\\input.xlsx";
        List<Student> students = new ArrayList<>();


        //Чтение данных из файла
        try (FileInputStream fis = new FileInputStream(inputFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            if (sheet == null) {
                logger.error("The sheet was not found in the input file.");
                throw new IOException("The sheet was not found in the file.");
            }
            for (Row row : sheet) {
                if (row.getRowNum() == 0 || row.getLastCellNum() < 2) continue;

                Cell nameCell = row.getCell(0);
                Cell gradeCell = row.getCell(1);

                if (nameCell == null || gradeCell == null || nameCell.getCellType() != CellType.STRING || gradeCell.getCellType() != CellType.NUMERIC) {
                    logger.error("Invalid data type in the string {}. Skip it.", row.getRowNum());
                    continue;
                }

                String name = nameCell.getStringCellValue();
                double grade = gradeCell.getNumericCellValue();
                students.add(new Student(name, grade));
            }

            String defaultOutputFileName = "grade_analysis_report.xlsx";


            Path outputFilePath = Paths.get(defaultOutputFileName);
            if (Files.exists(outputFilePath)) {
                logger.warn("Output file '{}' it already exists. Choose a different name or delete an existing file..", defaultOutputFileName);
                return;
            }
            // Анализ данных
            try (FileOutputStream outputStream = new FileOutputStream(defaultOutputFileName)) {
                analyzeGrades(students, outputStream, defaultOutputFileName);
                logger.info("The results of the analysis are recorded in the output file: {}", defaultOutputFileName);
                createChart(students);
            }

        } catch (IOException e) {
            logger.error("Error processing the file: {}", e.getMessage(), e); // Полная информация об ошибке
        } catch (Exception e) {
            logger.error("An error has occurred: {}", e.getMessage(), e);
        }
        logger.info("The application is shutting down");
    }

    /**
     * Analysis of student performance.
     *
     * @param students the list of students for whom the performance analysis was carried out
     * @param outputFileName the output file with the results of the analysis
     * @param outputStream stream to write to the output file
     * @throws IOException If an I/O error has occurred
     *
     *
     */

    private static void analyzeGrades(List<Student> students,FileOutputStream outputStream, String outputFileName) throws IOException {
        logger.info("Starting grade analysis");
        // Статистика
        int countExcellent = 0;
        int countGood = 0;
        int countSatisfactory = 0;
        int countFail = 0;
        double totalGrades = 0;
        double maxGrade = Double.MIN_VALUE;

        List<String> excellentStudents = new ArrayList<>();
        List<String> goodStudents = new ArrayList<>();
        List<String> satisfactoryStudents = new ArrayList<>();
        List<String> failedStudents = new ArrayList<>();

        try{
            for (Student student : students) {
                double grade = student.getGrade();
                totalGrades += grade;
                if (grade == EXCELLENT_GRADE) {
                    countExcellent++;
                    excellentStudents.add(student.getName());
                } else if (grade == GOOD_GRADE) {
                    countGood++;
                    goodStudents.add(student.getName());
                } else if (grade == SATISFACTORY_GRADE) {
                    countSatisfactory++;
                    satisfactoryStudents.add(student.getName());
                } else {
                    countFail++;
                    failedStudents.add(student.getName());
                }
                maxGrade = Math.max(maxGrade, grade);
            }
            logger.info("The analysis is completed, we are preparing to record the results");

            double averageGrade = students.isEmpty() ? 0 : totalGrades / students.size();
            // Запись в Excel файл (вынесена в отдельный метод)
            writeResultsToExcel(outputStream, outputFileName, countExcellent, countGood, countSatisfactory, countFail, maxGrade, averageGrade, excellentStudents, goodStudents, satisfactoryStudents, failedStudents);

            logger.info("The assessment analysis has been successfully completed");

        } catch (Exception e) {
            logger.error("An error occurred while analyzing the estimates: {}", e.getMessage(), e);
            throw new IOException("Error in analyzing estimates: " + e.getMessage(), e); // Перебрасываем исключение
        }
    }

    /**
     * Writing students to a new file.
     *
     * @param countExcellent the number of excellent students.
     * @param countGood the number of good students.
     * @param countSatisfactory the number of triples.
     * @param countFail the number of students not admitted in the group.
     * @param maxGrade the maximum score from the group.
     * @param averageGrade average score from the whole group.
     * @param excellentStudents the list of excellent students.
     * @param goodStudents the list of good students.
     * @param satisfactoryStudents yhe list of triples students.
     * @param failedStudents the list of not admitted students.
     * @param outputFileName the output file with the results of the analysis.
     * @param outputStream stream to write to the output file.
     * @throws IOException If an I/O error has occurred
     *
     */

    private static void writeResultsToExcel(FileOutputStream outputStream, String outputFileName, int countExcellent, int countGood, int countSatisfactory, int countFail, double maxGrade, double averageGrade, List<String> excellentStudents, List<String> goodStudents, List<String> satisfactoryStudents, List<String> failedStudents) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            logger.info("Starting to record the results in an Excel file: {}", outputFileName);
            Sheet sheet = workbook.createSheet("Analysis");
            CellStyle centeredStyle = workbook.createCellStyle();
            centeredStyle.setAlignment(HorizontalAlignment.CENTER);

            // Заголовки
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Критерий");
            int colNum = 1;
            addCategoryHeader(headerRow, colNum++, "Отличники", centeredStyle);
            addCategoryHeader(headerRow, colNum++, "Хорошисты", centeredStyle);
            addCategoryHeader(headerRow, colNum++, "Троечники", centeredStyle);
            addCategoryHeader(headerRow, colNum++, "Не допущенные", centeredStyle);
            headerRow.createCell(colNum).setCellValue("Максимальная оценка");

            // Количество
            Row countRow = sheet.createRow(1);
            countRow.createCell(0).setCellValue("Количество");
            colNum = 1;
            countRow.createCell(colNum++).setCellValue(countExcellent);
            countRow.createCell(colNum++).setCellValue(countGood);
            countRow.createCell(colNum++).setCellValue(countSatisfactory);
            countRow.createCell(colNum++).setCellValue(countFail);
            countRow.createCell(colNum).setCellValue(maxGrade);
            // Центрируем ячейки с количеством
            for (int i = 1; i < countRow.getLastCellNum(); i++) {
                countRow.getCell(i).setCellStyle(centeredStyle);
            }

            // Список студентов (в столбик)
            Row listRow = sheet.createRow(2);
            listRow.createCell(0).setCellValue("фио");
            int rowNum = 2;
            int maxRows = Math.max(Math.max(excellentStudents.size(), goodStudents.size()), Math.max(satisfactoryStudents.size(), failedStudents.size()));
            addStudentListToColumn(sheet, rowNum, 1, excellentStudents, maxRows);
            addStudentListToColumn(sheet, rowNum, 2, goodStudents, maxRows);
            addStudentListToColumn(sheet, rowNum, 3, satisfactoryStudents, maxRows);
            addStudentListToColumn(sheet, rowNum, 4, failedStudents, maxRows);


            // Средний балл
            Row averageRow = sheet.createRow(rowNum + maxRows + 1);
            averageRow.createCell(0).setCellValue("Средний балл группы");
            averageRow.createCell(1).setCellValue(averageGrade);
            averageRow.getCell(1).setCellStyle(centeredStyle);

            // Автоматическое изменение размера столбцов
            for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                sheet.autoSizeColumn(i);
            }
            try {
                // Запись данных, в try-catch для каждой записи
                workbook.write(outputStream);
                logger.info("The Excel file was successfully written: {}", outputFileName);
            } catch (IOException e) {
                logger.error("Error writing to Excel file: {}", outputFileName, e);
                throw e;
            }

        } catch (IOException e) {
            logger.error("Error when creating an Excel file: {}", e.getMessage(), e);
            throw new IOException("Error writing to Excel file: " + e.getMessage(), e);
        }
    }

    /**
     * Creating cells in the specified rows and columns and naming them.
     *
     * @param row the line in question.
     * @param colNum column number
     * @param category column names(categories).
     * @param centeredStyle a style for centering text.
     *
     *
     */

    private static void addCategoryHeader(Row row, int colNum, String category, CellStyle centeredStyle) {
        Cell cell = row.createCell(colNum);
        cell.setCellValue(category);
        cell.setCellStyle(centeredStyle);
    }

    /**
     *
     * Adding the full names of the students to the required columns and rows.
     *
     * @param sheet list.
     * @param startRow line number.
     * @param col column number.
     * @param names list of students of a specific category.
     * @param maxRows maximum number of rows.
     *
     *
     */

    private static void addStudentListToColumn(Sheet sheet, int startRow, int col, List<String> names, int maxRows) {
        try {
            for (int i = 0; i < names.size() || i < maxRows; i++) {
                Row row = sheet.getRow(startRow + i);
                if (row == null) {
                    row = sheet.createRow(startRow + i);
                }
                if(i < names.size()){
                    row.createCell(col).setCellValue(names.get(i));
                }else {
                    row.createCell(col).setCellValue("");
                }
            }
        } catch (Exception e){
            logger.error("Error when writing a list of students to a column {}: {}", col, e.getMessage(),e);
        }
    }

    /**
     * Creates a student progress chart based on the provided list of students.
     *
     * @param students the list of students for whom the schedule will be created.
     *
     */
    private static void createChart(List<Student> students) {
        try{
            DefaultCategoryDataset dataset = new DefaultCategoryDataset();
            dataset.addValue(students.size(), "Студенты", "Всего");
            dataset.addValue(countStudentsWithGrade(students, 5), "Отличники", "Отличники(5)");
            dataset.addValue(countStudentsWithGrade(students, 4), "Хорошисты", "Хорошисты(4)");
            dataset.addValue(countStudentsWithGrade(students, 3), "Троечники", "Троечники(3)");
            dataset.addValue(countStudentsWithGradesInRange(students), "Не допущены", "Не допущены(2 или 1)");

            JFreeChart barChart = ChartFactory.createBarChart(
                    "Статистика успеваемости студентов",
                    "Категории",
                    "Количество",
                    dataset,
                    PlotOrientation.VERTICAL,
                    true, true, false);

            // Визуализация графика
            SwingUtilities.invokeLater(() -> {
                JFrame frame = new JFrame();
                frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
                frame.add(new ChartPanel(barChart));
                frame.pack();
                frame.setLocationRelativeTo(null); // Центрируем окно на экране
                frame.setVisible(true);
            });
        }catch (Exception e){
            logger.error("Error when creating a graph: {}", e.getMessage(), e);

        }
    }

    /**
     * Counts the number of students with a given grade.
     *
     * @param students the list of students to be checked.
     * @param grade the assessment for which it is necessary to count the number of students.
     * @return the number of students with the specified grade.
     *
     */

    private static int countStudentsWithGrade(List<Student> students, double grade) {
        return (int) students.stream().filter(student -> student.getGrade() == grade).count();
    }
    /**
     * Counts the number of students with grades in the specified range.
     *
     * @param students the list of students to be checked.
     * @return the number of students with grades in the specified range.
     *
     */
    private static int countStudentsWithGradesInRange(List<Student> students) {
        return (int) students.stream().filter(student -> student.getGrade() >= 1 && student.getGrade() <= 2).count();
    }

    /**
     * A class to represent a student.
     */
    // Класс для представления студента
    static class Student {
        /** Student's name */
        private final String name;
        /** Student's assessment */
        private final double grade;

        /**
         * Constructor of the Student class
         * @param name Student's name
         * @param grade Student's assessment
         */
        public Student(String name, double grade) {
            if (name == null || name.trim().isEmpty()) {
                throw new IllegalArgumentException("The student's name cannot be empty.");
            }

            if (grade < 1 || grade > 5) {
                throw new IllegalArgumentException("The score should be in the range from 1 to 5.");
            }
            this.name = name;
            this.grade = grade;
        }

        /**
         * Gets the student's name
         * @return Student's name
         */
        public String getName() {
            return name;
        }

        /**
         * Gets a student's grade
         * @return Student's assessment
         */
        public double getGrade() {
            return grade;
        }
    }
}
