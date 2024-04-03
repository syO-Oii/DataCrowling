package com.crowling;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

public class PlayerStatusCrowling {
	
	private WebDriver driver;
	private WebElement element;
	private String url;
	private String[] team = {
			"1001", "삼성",
			"2002", "기아", 
			"3001",	"롯데",
			"5002",	"LG",
			"6002",	"두산",
			"7002",	"한화",
			"9002",	"SSG",
			"10001", "키움",
			"11001", "NC",
			"12001", "KT"
	};
	
 	// 1. 드라이버 설치 경로자급제
	public static String WEB_DRIVER_ID = "webdriver.chrome.driver";
	public static String WEB_DRIVER_PATH = "D:/chromedriver/chromedriver.exe";
	
	public PlayerStatusCrowling() {
		//WebDriver 경로 설정
		System.setProperty(WEB_DRIVER_ID, WEB_DRIVER_PATH);
		
		// 2. WebDriver 옵션 설정
		ChromeOptions options = new ChromeOptions();
		options.addArguments("--start-maximized");
		options.addArguments("--disable-popup-blocking");
		      
		driver = new ChromeDriver(options);
		
		// 크롤링 할 주소
		url = "https://statiz.sporki.com/stats/?m=main&m2=batting&m3=default&so=WAR&ob=DESC"
				+ "&year=2023&sy=&ey="	// year : 크롤링 할 연도
				+ "&te=2002&po="		// te : 팀코드
				+ "&lt=10100&reg=A&pe=&ds=&de=&we=&hr=&ha=&ct=&st=&vp=&bo=&pt=&pp=&ii=&vc=&um=&oo=&rr=&sc=&bc=&ba=&li=&as=&ae=&pl=&gc=&lr=&pr=50&ph=&hs=&us=&na=&ls=&sf1=&sk1=&sv1=&sf2=&sk2=&sv2=";
	}
	
	public void activateBot() {
		try {
			// 크롤링 주소 입력
			driver.get(url);
			Thread.sleep(2000); // 3. 페이지 로딩 대기 시간
			
			// 테이블 정보 저장
			WebElement table = driver.findElement(By.xpath("/html/body/div[2]/div[3]/section/div[8]"));
			
			// 새로운 Excel 워크북 생성
	        try (Workbook workbook = new XSSFWorkbook()) {
	            // Excel 시트 생성
	            Sheet sheet = workbook.createSheet("23기아");

	            // 테이블의 각 행(tr)을 가져와서 Excel 행(row)으로 생성
	            List<WebElement> rows = table.findElements(By.tagName("tr"));
	            int rowNum = 0;
	            for (WebElement row : rows) {
	                Row excelRow = sheet.createRow(rowNum++);
	                List<WebElement> cells = row.findElements(By.tagName("td"));
	                int cellNum = 0;
	                for (WebElement cell : cells) {
	                    // Excel 셀 생성 및 값 설정
	                    Cell excelCell = excelRow.createCell(cellNum++);
	                    excelCell.setCellValue(cell.getText());
	                }
	            }

	            // Excel 파일로 저장
	            try (FileOutputStream outputStream = new FileOutputStream("table_data.xlsx")) {
	                workbook.write(outputStream);
	            }
	            System.out.println("Excel 파일이 생성되었습니다.");
	        } catch (IOException e) {
	            e.printStackTrace();
	        }
		}catch (Exception e) {
			e.printStackTrace();
		} finally {
			driver.close(); // 5. 브라우저 종료
		}
	}
	
	
	public static void main(String[] args) {
		
		PlayerStatusCrowling psc = new PlayerStatusCrowling();
		psc.activateBot();
		
		
	}

}
