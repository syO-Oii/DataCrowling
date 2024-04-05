package com.crowling;

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

public class HitterStatusCrowling {
	
	private WebDriver driver;
	private String hitterUrl;
	
	
 	// 1. 드라이버 설치 경로
	public static String WEB_DRIVER_ID = "webdriver.chrome.driver";
	public static String WEB_DRIVER_PATH = "D:/chromedriver/chromedriver.exe";
	
	public HitterStatusCrowling() {
		//WebDriver 경로 설정
		System.setProperty(WEB_DRIVER_ID, WEB_DRIVER_PATH);
		
		// 2. WebDriver 옵션 설정
		ChromeOptions options = new ChromeOptions();
		options.addArguments("--start-maximized");
		options.addArguments("--disable-popup-blocking");
		      
		driver = new ChromeDriver(options);
		
		// 크롤링 할 주소
		hitterUrl = "https://statiz.sporki.com/stats/?m=main&"
			+ "m2=batting"	// 타자
			+ "&m3=default&so=WAR&ob=DESC&"
			+ "year=2023"	// 연도
			+ "&sy=&ey=&te=&po=&lt=10100&"
			+ "reg=A"		// 규정이닝 : 전체
			+ "&pe=&ds=&de=&we=&hr=&ha=&ct=&st=&vp=&bo=&pt=&pp=&ii=&vc=&um=&oo=&rr=&sc=&bc=&ba=&li=&as=&ae=&pl=&gc=&lr=&"
			+ "pr=500"		// 출력 수
			+ "&ph=&hs=&us=&na=&ls=0&sf1=G&sk1=&sv1=&sf2=G&sk2=&sv2=-25";
	}
	
	public void activateBot() {
		try {
			// 크롤링 주소 입력
			driver.get(hitterUrl);
			Thread.sleep(2000); // 3. 페이지 로딩 대기 시간
			
			// 테이블 정보 저장
			WebElement table = driver.findElement(By.xpath("/html/body/div[2]/div[3]/section/div[8]"));
			
			// 새로운 Excel 워크북 생성
	        try (Workbook workbook = new XSSFWorkbook()) {
	            // Excel 시트 생성
	            Sheet sheet = workbook.createSheet("23전체타자");

	            // 테이블의 각 행(tr)을 가져와서 Excel 행(row)으로 생성
	            List<WebElement> rows = table.findElements(By.tagName("tr"));
	            int rowNum = 0;
	            
	            // 팀 코드와 팀 이름 매핑을 위한 Map 생성
                Map<String, String> teamMap = new HashMap<>();
                teamMap.put("1001", "삼성");
                teamMap.put("2002", "기아");
                teamMap.put("3001", "롯데");
                teamMap.put("5002", "LG");
                teamMap.put("6002", "두산");
                teamMap.put("7002", "한화");
                teamMap.put("9002", "SSG");
                teamMap.put("10001", "키움");
                teamMap.put("11001", "NC");
                teamMap.put("12001", "KT");
                
                // 첫 번째 행에 열의 이름 추가
                Row headerRow = sheet.createRow(rowNum++);
                String[] headers = {"rank", "p_no", "name", "year", "position", "team", "WAR", "G", "PA", "ePA", "AB", "R", "H", "2B", "3B", "HR", "TB", "RBI", "SB", "CS", "BB", "HP", "IB", "SO", "GDP", "SH", "SF", "AVG", "OBP", "SLG", "OPS", "R/ePA", "wRC+", "WAR"};
                for (int i = 0; i < headers.length; i++) {
                    Cell cell = headerRow.createCell(i);
                    cell.setCellValue(headers[i]);
                }

	            for (int i = 0; i < rows.size(); i++) {
	                // 12의 배수인 경우 해당 행 생략
	                if (i % 12 == 0 || i % 12 == 1) {
	                    continue;
	                }
	                
	                WebElement row = rows.get(i);
	             // 각 행의 데이터를 가져옴
	                List<WebElement> cells = row.findElements(By.tagName("td"));
	                // Excel 행(row) 생성
	                Row excelRow = sheet.createRow(rowNum++);
	                int cellNum = 0;
	                
	                for (WebElement cell : cells) {
	                    // Excel 셀 생성 및 값 설정
	                    Cell excelCell = excelRow.createCell(cellNum++);
	                    String cellValue = cell.getText().trim();
	                    
	                    // 선수 고유번호 추출 후 셀정보 삽입
	                    if(cellNum == 2) {
	                    	WebElement anchorElement = cell.findElement(By.tagName("a"));
	                        String anchorHref = anchorElement.getAttribute("href");
	                        String[] parts = anchorHref.split("&");
	                        for (String part : parts) {
	                            if (part.contains("p_no")) {
	                                String p_no = part.split("=")[1];
	                                excelCell.setCellValue(p_no);
	                                excelCell = excelRow.createCell(cellNum++);
	                                excelCell.setCellValue(cellValue);
	                                break;
	                            }
	                        }
	                        
	                    // 팀정보, 연도정보 추출 후 셀정보 삽입
	                    } else if (cellNum == 4) { // 팀 이미지 정보는 네 번째 열(td)에 있음
	                    	cellNum--;
	                    	WebElement spanElement1 = cell.findElement(By.xpath(".//span[1]"));
	                        WebElement spanElement2 = cell.findElement(By.xpath(".//span[3]"));
	                        String spanText1 = spanElement1.getText();
	                        String spanText2 = spanElement2.getText();

	                        WebElement imgElement = cell.findElement(By.tagName("img"));
	                        String srcValue = imgElement.getAttribute("src");
	                        String teamCode = srcValue.substring(srcValue.lastIndexOf("/") + 1, srcValue.lastIndexOf("."));
	                        
	                        // Map에서 팀 이름 가져오기
                            String teamName = teamMap.get(teamCode);
	                        
	                        // 각 span 요소의 내용을 셀에 저장
	                        Cell spanCell1 = excelRow.createCell(cellNum++);
	                        spanCell1.setCellValue(spanText1);
	                        Cell spanCell2 = excelRow.createCell(cellNum++);
	                        spanCell2.setCellValue(spanText2);
	                        Cell teamImgCell = excelRow.createCell(cellNum++);
	                        teamImgCell.setCellValue(teamName);
	                    } else {	// 특이사항이 없다면 셀 정보 삽입
	                    	excelCell.setCellValue(cellValue);
	                    }
	                }          
	            }
	            // Excel 파일로 저장
	            try (FileOutputStream outputStream = new FileOutputStream("2023_Hitter_Status.xlsx")) {
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
		
		HitterStatusCrowling psc = new HitterStatusCrowling();
		psc.activateBot();
		
		
	}

}