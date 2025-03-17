package com.bbstzb.analyze.util;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * 统计数据生成excel
 */
public class DynamicExcelProcessor {

    static class CharMapping {
        String chineseChar;
        String correspondingValue;

        CharMapping(String chineseChar, String correspondingValue) {
            this.chineseChar = chineseChar;
            this.correspondingValue = correspondingValue;
        }
    }

    private boolean getExcelInfo(String inputPath, String outputPath) {
        boolean flag = false;
        List<CharMapping> mappings = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(inputPath);
             Workbook inputWorkbook = WorkbookFactory.create(fis)) {

            Sheet sheet = inputWorkbook.getSheetAt(0);
            final int CHINESE_COL = 1;  // 中文列固定为B列（索引1）

            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue; // 跳过标题行

                Cell chineseCell = row.getCell(CHINESE_COL);
                if (chineseCell == null || chineseCell.getCellType() != CellType.STRING) continue;

                String chinese = chineseCell.getStringCellValue().trim();
                int charCount = chinese.length();
                if (charCount == 0) continue;

                // 动态计算起始列（C列开始，索引2）
                int startDataCol = 2;

                for (int i = 0; i < charCount; i++) {
                    int targetCol = startDataCol + i;
                    Cell dataCell = row.getCell(targetCol);

                    if (dataCell != null) {
                        String charStr = chinese.substring(i, i + 1);
                        String value = getCellValue(dataCell);
                        mappings.add(new CharMapping(charStr, value));
                    }
                }
            }

            generateOutput(mappings, outputPath);
             flag = true;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return flag;
    }

    private static void generateOutput(List<CharMapping> mappings, String path) {
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(path)) {
            Sheet sheet = workbook.createSheet("字符映射");

            // 创建标题
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("中文字符");
            header.createCell(1).setCellValue("对应值");
            // 填充数据
            int rowNum = 1;
            for (CharMapping mapping : mappings) {
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(mapping.chineseChar);
                row.createCell(1).setCellValue(mapping.correspondingValue);
            }
            workbook.write(fos);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static String getCellValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                return String.valueOf((int) cell.getNumericCellValue());
            default:
                return "";
        }
    }
}
