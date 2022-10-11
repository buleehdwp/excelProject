package com.excel;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.hibernate.annotations.common.util.impl.LoggerFactory;
import org.jboss.logging.Logger;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.List;

public class Excel {
    private final Logger log = LoggerFactory.logger(Excel.class);
    //  , , , keyword5 = 식자재,
    private final String[] keywords1 = {"한끼마라탕", "맛맛향", "박리분식", "표표마라탕", "빵굼터"};// keywords1 = 식비,
    private final String[] keywords2 = {"GS25", "씨유", "세븐일레븐", "이마트24"};// keywords2 = 편의점
    private final String[] keywords3 = {"쿠팡", "슬기로운 간식생활", "주식회사 미로", "(주) 브랜디", "씨엔티테크_카카오페이"};// keywords3 = 쇼핑(쿠팡, 다이소)
    private final String[] keywords4 = {"쿠팡이츠"};// keywords4 = 배달음식
    private final String[] keywords5 = {"온마켓", "삼중정육점"};// keywords5 = 식자재
    private final String[] keywords6 = {"네이버페이", "코인", "SKT요금", "JetBrains", "페이타랩", "(주)그린카"};// keywords6 = 여가생활(데이트, 친구들, 커피, 통신비)
    private final String[] keywords7 = {"버스", "지하철"};// keywords6 = 교통비

    public static void main(String[] args) {
        try {
            FileInputStream fis = new FileInputStream("C:\\oldExcel.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            Excel excel = new Excel();
            //            excel.rebuildExcel(workbook); // 엑셀 날짜별 row 정리
            excel.excelRebuild(workbook.getSheetAt(0));

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void excelLineBreak(XSSFWorkbook workbook) throws Exception {
        // 2번째 row부터 데이터 추출
        // cell_0 = 결제일, cell_2 = 카드구분, cell_3 = 결제장소, cell_4 = 금액
        XSSFSheet sheet = workbook.getSheetAt(0);// 1번째 시트
        int rows = sheet.getPhysicalNumberOfRows();

        for (int i = 2; i < rows - 1; i++) {
            XSSFRow targetRow = sheet.getRow(i);
            if (targetRow == null) {
                continue;
            }

            XSSFRow compareRow = sheet.getRow(i + 1);
            String targetCell = targetRow.getCell(0).getStringCellValue(); // 비교 1 결제일
            String compareCell = compareRow.getCell(0).getStringCellValue(); // 비교 2 결제일

            if (targetCell.equals(compareCell)) { // 비교1 != 비교2 이면 1 row  삽입
                continue;
            } else {
                sheet.shiftRows(i + 1, rows, 1); // 1 row 삽입
                rows++; // 전체 row 수 증가
            }

        }
        FileOutputStream fos = new FileOutputStream("C:\\excelLineBreak.xlsx");
        workbook.write(fos);
    }

    public void excelRebuild(XSSFSheet sourceSheet) throws Exception {
        XSSFWorkbook rebuildWorkbook = new XSSFWorkbook();
        // 시트 지정
        Calendar calendar = Calendar.getInstance();
        int month = calendar.get(Calendar.MONTH) + 1;
        XSSFSheet sheet = rebuildWorkbook.createSheet(month + "월");

        // 양식
        createTemplates(sheet);

        // 가계부 작성
        createAccountBook(sheet, sourceSheet);


        FileOutputStream fos = new FileOutputStream("C:\\excelRebuild.xlsx");
        rebuildWorkbook.write(fos);
    }

    // 상단 양식
    public void createTemplates(XSSFSheet sheet) {
        XSSFRow row = sheet.createRow(1);
        row.createCell(1).setCellValue("날짜");

        row.createCell(2).setCellValue("식비");
        row.createCell(3);

        row.createCell(4).setCellValue("편의점");
        row.createCell(5);

        row.createCell(6).setCellValue("쇼핑(쿠팡, 다이소)");
        row.createCell(7);

        row.createCell(8).setCellValue("배달음식");
        row.createCell(9);

        row.createCell(10).setCellValue("식자재");
        row.createCell(11);

        row.createCell(12).setCellValue("여가생활(데이트, 커피, 통신비, 친구들)");
        row.createCell(13);

        row.createCell(14).setCellValue("교통비");
        row.createCell(15);

        row.createCell(16).setCellValue("미분류");
        row.createCell(17);

        sheet.addMergedRegion(new CellRangeAddress(2, 2, 2, 3));
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 4, 5));
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 6, 7));
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 8, 9));
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 10, 11));
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 12, 13));
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 14, 15));
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 16, 17));
    }

    // 가계부 작성
    public void createAccountBook(XSSFSheet sheet, XSSFSheet sourceSheet) {
        int rows = sourceSheet.getPhysicalNumberOfRows() - 1;
        double transportationCost = 0; // 교통비
        int createRowNum = 2;

        for (int i = 2; i <= rows; i++) { // 명세서 row 수만큼
            String targetDate = sourceSheet.getRow(i).getCell(0).getStringCellValue(); // 날짜
            String targetName = sourceSheet.getRow(i).getCell(3).getStringCellValue(); // 결제 내역
            double targetMoney = sourceSheet.getRow(i).getCell(4).getNumericCellValue(); // 결제 금액

            Boolean flag2 = false; // 미분류 확인, 분류시 true
            String compareDate = (sourceSheet.getRow(i + 1) != null) ? sourceSheet.getRow(i + 1).getCell(0).getStringCellValue() : "";

            // 식비
            for (String keyword : keywords1) {
                if (targetName.contains(keyword)) {
                    XSSFRow row = sheet.createRow(createRowNum++); // 추가할 row
                    row.createCell(1).setCellValue(targetDate);
                    row.createCell(2).setCellValue(targetName);
                    row.createCell(3).setCellValue(targetMoney);
                    flag2 = true;
                    break;
                }
            }


            // 편의점
            for (String keyword : keywords2) {
                if (targetName.contains(keyword)) {
                    XSSFRow row = sheet.createRow(createRowNum++); // 추가할 row
                    row.createCell(1).setCellValue(targetDate);
                    row.createCell(4).setCellValue(targetName);
                    row.createCell(5).setCellValue(targetMoney);
                    flag2 = true;
                    break;
                }
            }

            // 쇼핑(쿠팡, 다이소)
            for (String keyword : keywords3) {
                if (!targetName.equals("쿠팡이츠") && targetName.contains(keyword)) {
                    XSSFRow row = sheet.createRow(createRowNum++); // 추가할 row
                    row.createCell(1).setCellValue(targetDate);
                    row.createCell(6).setCellValue(targetName);
                    row.createCell(7).setCellValue(targetMoney);
                    flag2 = true;
                    break;
                }
            }

            // 배달음식
            for (String keyword : keywords4) {
                if (!targetName.equals("쿠팡") && targetName.contains(keyword)) {
                    XSSFRow row = sheet.createRow(createRowNum++); // 추가할 row
                    row.createCell(1).setCellValue(targetDate);
                    row.createCell(8).setCellValue(targetName);
                    row.createCell(9).setCellValue(targetMoney);
                    flag2 = true;
                    break;
                }
            }

            // 식자재
            for (String keyword : keywords5) {
                if (targetName.contains(keyword)) {
                    XSSFRow row = sheet.createRow(createRowNum++); // 추가할 row
                    row.createCell(1).setCellValue(targetDate);
                    row.createCell(10).setCellValue(targetName);
                    row.createCell(11).setCellValue(targetMoney);
                    flag2 = true;
                    break;
                }
            }

            // 여가생활(데이트, 친구들, 커피, 통신비)
            for (String keyword : keywords6) {
                if (targetName.contains(keyword)) {
                    XSSFRow row = sheet.createRow(createRowNum++); // 추가할 row
                    row.createCell(1).setCellValue(targetDate);
                    row.createCell(12).setCellValue(targetName);
                    row.createCell(13).setCellValue(targetMoney);
                    flag2 = true;
                    break;
                }
            }

            // 교통비 합산
            for (String keyword : keywords7) {
                if (targetName.contains(keyword)) {
                    createRowNum--;
                    transportationCost += targetMoney;
                    flag2 = true;
                    break;
                }
            }

            // 미분류
            if (!flag2) {
                XSSFRow row = sheet.createRow(createRowNum++); // 추가할 row
                row.createCell(1).setCellValue(targetDate);
                row.createCell(16).setCellValue(targetName);
                row.createCell(17).setCellValue(targetMoney);
            }
        }
        sheet.getRow(createRowNum - 1).createCell(14).setCellValue("교통비 통합");
        sheet.getRow(createRowNum - 1).createCell(15).setCellValue(transportationCost);
    }
}
