package com.crowling;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

public class PitcherStatusCrowling_KBO_paging {
    
    private WebDriver driver;
    private String hitterUrl;
    
    // 새로운 엑셀 파일 경로
    private String newExcelFilePath = "2024_KBO_Pitcher_Status.xlsx";

    // 1. 드라이버 설치 경로
    public static String WEB_DRIVER_ID = "webdriver.chrome.driver";
    public static String WEB_DRIVER_PATH = "D:/chromedriver/chromedriver.exe";
    
    // 현재 페이지의 XPath
    private String currentPageXPath = "//*[@id=\"cphContents_cphContents_cphContents_ucPager_btnNo1\"]";

    // 다음 페이지로 이동하는 버튼의 XPath
    private String nextPageXPath = "//*[@id=\"cphContents_cphContents_cphContents_ucPager_btnNo";

    public PitcherStatusCrowling_KBO_paging() {
        // WebDriver 경로 설정
        System.setProperty(WEB_DRIVER_ID, WEB_DRIVER_PATH);
        
        // WebDriver 옵션 설정
        ChromeOptions options = new ChromeOptions();
        options.addArguments("--start-maximized");
        options.addArguments("--disable-popup-blocking");
              
        driver = new ChromeDriver(options);
        
        // 크롤링 할 주소
        hitterUrl = "https://www.koreabaseball.com/Record/Player/PitcherBasic/Basic1.aspx";
    }
    
    public void activateBot() {
        try {
            // 크롤링 주소 입력
            driver.get(hitterUrl);
            Thread.sleep(2000); // 페이지 로딩 대기 시간
            
            // 페이지 수집 시작
            collectDataWithPaging();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            driver.quit(); // 브라우저 종료
        }
    }

    // 페이지당 데이터 수집하는 메소드
    public void collectDataOnPage() {
        // 테이블 정보 저장
        WebElement table = driver.findElement(By.xpath("//*[@id=\"cphContents_cphContents_cphContents_udpContent\"]/div[3]/table/tbody"));
        
        // 기존 엑셀 파일 확인 및 불러오기
        Workbook workbook;
        Sheet sheet;
        File excelFile = new File(newExcelFilePath);
        
        if (excelFile.exists()) {
            try {
                FileInputStream fis = new FileInputStream(excelFile);
                workbook = new XSSFWorkbook(fis);
                sheet = workbook.getSheetAt(0); // 첫 번째 시트 가져오기
            } catch (IOException e) {
                e.printStackTrace();
                return;
            }
        } else {
            workbook = new XSSFWorkbook();
            sheet = workbook.createSheet("24전체타자");
        }
        
        // Excel 시트에서 마지막 행 번호 가져오기
        int lastRowNum = sheet.getLastRowNum();
        
        // 데이터 수집 및 기존 엑셀 파일에 추가
        collectAndAppendData(table, sheet, lastRowNum);
        
        // 파일 저장
        try (FileOutputStream fos = new FileOutputStream(newExcelFilePath)) {
            workbook.write(fos);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // 페이징을 통한 전체 데이터 수집
    public void collectDataWithPaging() {
        int pageCounter = 1;
        while (true) {
            try {
                pageCounter++;
                // 현재 페이지의 데이터 수집
                collectDataOnPage();
                if (pageCounter >= 6) break; // 3 페이지까지만 크롤링
                
                // 다음 페이지로 이동
                WebElement nextPageButton = driver.findElement(By.xpath(nextPageXPath+pageCounter+"\"]"));
                nextPageButton.click();
                
                // 페이지 로딩 대기 시간
                Thread.sleep(2000);
            } catch (NoSuchElementException e) {
                // 다음 페이지 버튼을 찾을 수 없으면 페이징 종료
                System.out.println("더 이상 다음 페이지가 없습니다.");
                break;
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        }
    }
    
    // 데이터 수집 및 기존 엑셀 파일에 추가
    public void collectAndAppendData(WebElement table, Sheet sheet, int lastRowNum) {
        // 테이블의 각 행(tr)을 가져와서 Excel 행(row)으로 생성
        List<WebElement> rows = table.findElements(By.tagName("tr"));
        
        for (int i = 0; i < rows.size(); i++) {
            WebElement row = rows.get(i);
            List<WebElement> cells = row.findElements(By.tagName("td"));
            Row excelRow = sheet.createRow(++lastRowNum); // 행 번호를 증가시키면서 추가
            
            int cellNum = 0;
            for (WebElement cell : cells) {
                Cell excelCell = excelRow.createCell(cellNum++);
                String cellValue = cell.getText().trim();
                
                if (cellNum == 2) { // 선수 고유번호 추출 후 셀 정보 삽입
                    WebElement anchorElement = cell.findElement(By.tagName("a"));
                    String anchorHref = anchorElement.getAttribute("href");
                    
                    // playerId 값을 추출하기 위해 href 속성값에서 'playerId=' 부분을 찾음
                    int startIndex = anchorHref.indexOf("playerId=") + 9; // 'playerId='의 인덱스 위치 + 9 (길이)
                    String playerId = anchorHref.substring(startIndex);
                    
                    excelCell.setCellValue(playerId);
                    excelCell = excelRow.createCell(cellNum++);
                    excelCell.setCellValue(cellValue);
                } else { // 특이사항이 없다면 셀 정보 삽입
                    excelCell.setCellValue(cellValue);
                }
            }
        }
    }
    
    public static void main(String[] args) {
        PitcherStatusCrowling_KBO_paging psc = new PitcherStatusCrowling_KBO_paging();
        psc.activateBot();
    }
}