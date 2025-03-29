package com.bbstzb.analyze.util;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class ExcelAnalysis {

    public static void main(String[] args) {
        try {
            String inputPath = "/Users/bajiu/Documents/output.xlsx";
            String statsPath = "/Users/bajiu/Documents/final_merged_statsnew.xlsx";
            File inputFile = new File(inputPath); // 输入Excel文件路径
            File outputFile = new File(statsPath); // 输出Excel文件路径
            FileInputStream fis = new FileInputStream(inputFile);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0);

            Map<String, Set<String>> dataMap = new LinkedHashMap<>();

            // 从第二行开始读取（索引1）
            for (int i = 1; i <= sheet.getPhysicalNumberOfRows(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    // 获取第一列和第二列的单元格
                    Cell chineseCell = row.getCell(0);
                    Cell readingCell = row.getCell(1);

                    // 如果单元格不为空，处理该行数据
                    if (chineseCell != null && readingCell != null) {
                        String chineseCharacter = chineseCell.getStringCellValue().trim();
                        String reading = readingCell.getStringCellValue().trim();

                        // 如果汉字已经存在于map中，则添加读音
                        dataMap.computeIfAbsent(chineseCharacter, k -> new HashSet<>()).add(reading);
                    }
                }
            }

            // 创建新的工作簿和工作表
            Workbook outputWorkbook = new XSSFWorkbook();
            Sheet outputSheet = outputWorkbook.createSheet("分析结果");

            // 写入表头
            Row headerRow = outputSheet.createRow(0);
            headerRow.createCell(0).setCellValue("汉字");
            headerRow.createCell(1).setCellValue("标音");
            headerRow.createCell(2).setCellValue("频次");

            // 写入数据
            int rowIndex = 1;
            for (Map.Entry<String, Set<String>> entry : dataMap.entrySet()) {
                Row row = outputSheet.createRow(rowIndex++);
                String chineseCharacter = entry.getKey();
                Set<String> readings = entry.getValue();

                // 将汉字写入第一个单元格，并合并该单元格
                Cell chineseCell = row.createCell(0);
                chineseCell.setCellValue(chineseCharacter);

                // 只有当当前行不为最后一行时，才合并单元格
                if (rowIndex > 1) {
                    //outputSheet.addMergedRegion(new CellRangeAddress(rowIndex - 1, rowIndex - 1, 0, 0));
                    CellStyle style = outputWorkbook.createCellStyle();
                    style.setAlignment(HorizontalAlignment.CENTER); // 居中对齐
                    chineseCell.setCellStyle(style);
                }

                // 将多个读音显示在第二列
                row.createCell(1).setCellValue(String.join(", ", readings));
                row.createCell(2).setCellValue(readings.size()); // 不同读音的频次
            }

            // 将数据写入到新的文件
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                outputWorkbook.write(fos);
            }

            // 关闭资源
            fis.close();
            workbook.close();
            outputWorkbook.close();

            System.out.println("Excel文件已成功处理并保存。");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
