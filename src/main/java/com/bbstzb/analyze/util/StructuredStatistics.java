package com.bbstzb.analyze.util;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;

public class StructuredStatistics {

    static class CharRecord {
        String chinese;
        String value;
        int count;
        int total;

        CharRecord(String chinese, String value, int count, int total) {
            this.chinese = chinese;
            this.value = value;
            this.count = count;
            this.total = total;
        }
    }

    private boolean getWords(String inputPath, String statsPath) {
        boolean flag = false;
        Map<String, Map<String, Integer>> statsMap = new TreeMap<>();

        try (FileInputStream fis = new FileInputStream(inputPath);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue;

                String chinese = getCellValue(row.getCell(0));
                String value = getCellValue(row.getCell(1));

                if (!chinese.isEmpty() && !value.isEmpty()) {
                    statsMap.computeIfAbsent(chinese, k -> new HashMap<>())
                            .merge(value, 1, Integer::sum);
                }
            }

            generateFinalMergedFile(statsMap, statsPath);
            flag = true;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return flag;
    }

    private static void generateFinalMergedFile(Map<String, Map<String, Integer>> statsMap, String path) {

        List<CharRecord> records = new ArrayList<>();

        // 转换数据结构
        for (Map.Entry<String, Map<String, Integer>> entry : statsMap.entrySet()) {
            String chinese = entry.getKey();
            Map<String, Integer> valueCounts = entry.getValue();
            int total = valueCounts.values().stream().mapToInt(Integer::intValue).sum();

            // 添加具体值统计
            for (Map.Entry<String, Integer> e : valueCounts.entrySet()) {
                records.add(new CharRecord(chinese, e.getKey(), e.getValue(), total));
            }
        }

        // 写入Excel
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(path)) {

            Sheet sheet = workbook.createSheet("统计结果");
            createHeader(sheet);

            int rowNum = 1;
            String lastChinese = null;
            int mergeStart = 1;

            for (int i = 0; i < records.size(); i++) {
                CharRecord record = records.get(i);
                Row row = sheet.createRow(rowNum++);

                // 汉字列和总计列合并逻辑
                if (!record.chinese.equals(lastChinese)) {
                    if (lastChinese != null && rowNum - 1 > mergeStart) {
                        // 合并上一个汉字的单元格（仅当有多个内容时）
                        sheet.addMergedRegion(new CellRangeAddress(
                                mergeStart, rowNum - 1, 0, 0)); // 汉字列合并
                        sheet.addMergedRegion(new CellRangeAddress(
                                mergeStart, rowNum - 1, 3, 3)); // 总计列合并
                    }
                    mergeStart = rowNum;
                    lastChinese = record.chinese;
                }

                // 填充数据
                row.createCell(0).setCellValue(record.chinese);
                row.createCell(1).setCellValue(record.value);
                row.createCell(2).setCellValue(record.count);
                row.createCell(3).setCellValue(record.total);
            }

            // 合并最后一个汉字的单元格（仅当有多个内容时）
            if (lastChinese != null && rowNum - 1 > mergeStart) {
                sheet.addMergedRegion(new CellRangeAddress(
                        mergeStart, rowNum - 1, 0, 0)); // 汉字列合并
                sheet.addMergedRegion(new CellRangeAddress(
                        mergeStart, rowNum - 1, 3, 3)); // 总计列合并
            }

            // 设置样式
            CellStyle totalStyle = workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setBold(true);
            totalStyle.setFont(font);

            // 标记总计列
            for (int i = 1; i < rowNum; i++) {
                Row row = sheet.getRow(i);
                row.getCell(3).setCellStyle(totalStyle);
            }

            // 自动调整列宽
            for (int i = 0; i < 4; i++) {
                sheet.autoSizeColumn(i);
            }

            workbook.write(fos);
            System.out.println("最终合并统计文件生成成功: " + path);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void createHeader(Sheet sheet) {
        Row header = sheet.createRow(0);
        String[] headers = {"汉字", "对应内容", "出现次数", "总计"};
        for (int i = 0; i < headers.length; i++) {
            header.createCell(i).setCellValue(headers[i]);
        }
    }

    private static String getCellValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue().trim();
            case NUMERIC: return String.valueOf((int) cell.getNumericCellValue());
            default: return "";
        }
    }
}
