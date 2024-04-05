package com.crowling;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

public class ProfileCrowling {
	String loadExcelFilePath = "2023_allPlayerInfo.xlsx";
	private WebDriver driver;
	// 1. 드라이버 설치 경로
	public static String WEB_DRIVER_ID = "webdriver.chrome.driver";
	public static String WEB_DRIVER_PATH = "D:/chromedriver/chromedriver.exe";

	public ProfileCrowling() {
		// WebDriver 경로 설정
		System.setProperty(WEB_DRIVER_ID, WEB_DRIVER_PATH);

		// 2. WebDriver 옵션 설정
		ChromeOptions options = new ChromeOptions();
		options.addArguments("--start-maximized");
		options.addArguments("--disable-popup-blocking");

		driver = new ChromeDriver(options);
	}

	public void activateBot() {
		// 파일 읽기
		try (FileInputStream excelFile = new FileInputStream(loadExcelFilePath);
				Workbook workbook = new XSSFWorkbook(excelFile);
				FileOutputStream outFile = new FileOutputStream("2023_allPlayerInfoPlus")) {
			
			
			
			// 첫 번째 시트에서 헤더 행 읽기
			Sheet sheet = workbook.getSheetAt(0);
			Row headerRow = sheet.getRow(0);
			headerRow.createCell(6).setCellValue("birth");
            headerRow.createCell(7).setCellValue("total");
                  
			// "p_no" 열의 인덱스 찾기
			int pNoColumnIndex = -1;
			for (Cell cell : headerRow) {
				if (cell.getStringCellValue().equals("p_no")) {
					pNoColumnIndex = cell.getColumnIndex();
					break;
				}
			}

			if (pNoColumnIndex == -1) {
				System.err.println("헤더에서 'p_no' 열을 찾을 수 없습니다.");
				return;
			}

			// 데이터 크롤링
			for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
				Row row = sheet.getRow(rowIndex);

				// 현재 행에서 p_no 값을 읽어옴
				Cell pNoCell = row.getCell(pNoColumnIndex);

				if (pNoCell != null && pNoCell.getCellType() == CellType.STRING) {
					int p_no = Integer.parseInt(pNoCell.getStringCellValue());

					// 웹 페이지에 접속하여 데이터 크롤링
					String url = "https://statiz.sporki.com/player/?m=playerinfo&p_no=" + p_no;
					driver.get(url);
					Thread.sleep(1000); // 3. 페이지 로딩 대기 시간

					// 크롤링할 데이터 요소 선택
					WebElement birthElement = driver
							.findElement(By.xpath("/html/body/div[2]/div[3]/section/div[3]/div[1]/ul/li[1]"));
					WebElement totaElement = driver.findElement(
							By.xpath("/html/body/div[2]/div[3]/section/div[3]/div[1]/div[3]/div[2]/span[3]"));
					String birth = birthElement.getText();
					String tota = totaElement.getText();
					// 데이터 출력
					birth = birth.replace("생년월일 :", "").trim();
					
					Cell birthCell = row.createCell(6);
					Cell totaCell = row.createCell(7);
					
					birthCell.setCellValue(birth);
					totaCell.setCellValue(tota);
				}
			}
			workbook.write(outFile);
		} catch (IOException e) {
			e.printStackTrace();
		} catch (InterruptedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			try {
				driver.quit(); // WebDriver 종료
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
	}

	public static void main(String[] args) {

		ProfileCrowling psc = new ProfileCrowling();
		psc.activateBot();

	}

}