package com.bbstzb.analyze.controller;

import com.bbstzb.analyze.util.AnalyzeUtil;
import jakarta.servlet.ServletOutputStream;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;


import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import org.springframework.core.io.Resource;
import org.springframework.http.ResponseEntity;

@RestController
@CrossOrigin(origins = "http://localhost:8080")
public class FileUploadController {

    @PostMapping("/upload")
    public void  handleFileUpload(@RequestParam("file") MultipartFile file, HttpServletResponse response) {
        try (XWPFDocument document = new XWPFDocument(file.getInputStream())) {
            List<XWPFParagraph> paragraphs = document.getParagraphs();
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Sentences");
            int rowCount = 0;
            for (XWPFParagraph para : paragraphs) {
                String[] sentences = para.getText().split("。");
                for (String sentence : sentences) {
                    boolean result = AnalyzeUtil.isComplexSentence(sentence);
                    Row row = sheet.createRow(rowCount);
                    Cell cell = row.createCell(0);
                    cell.setCellValue(sentence.trim());
                    Cell cell1 = row.createCell(1);
                    cell1.setCellValue(result);
                }
            }
                try {
                    response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=GBK");
                    response.setHeader("Content-disposition", "attachment; filename=test.xlsx");
                    ServletOutputStream out = null;
                    out = response.getOutputStream();
                    workbook.write(out);
                    out.flush();
                    out.close();
                }catch(Exception e){
                    e.printStackTrace();
                }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
