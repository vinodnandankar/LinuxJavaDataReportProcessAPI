package com.example.demo.service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Bean;
import org.springframework.core.ParameterizedTypeReference;
import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpMethod;
import org.springframework.http.MediaType;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Service;
import org.springframework.web.client.RestTemplate;

import com.example.demo.entity.LinuxJavaDataSheet;

@Service
public class DataReportPApiService {
	@Value("${excel.filepath}")
	private String filepath;
	@Value("${date.format}")
	private String dateFormat;
	public static final Integer COLUMN_COMMENT_NUM = 27;
	public static final Integer COLUMN_PROPOSED_DATE_NUM = 28;
	
	@Bean
	public RestTemplate restTemplate() {
		return new RestTemplate();
	}
	
	
	/*-----------------------Code to read System Api---------------------------*/
	@Scheduled(fixedRate = 1000)
	public Map<String, Object> getDataReport() {
		System.out.println("Executing Schedular for second======================================");
		Map<String, Object> response = new HashMap<>();

		HttpHeaders headers = new HttpHeaders();
		headers.setAccept(Arrays.asList(MediaType.APPLICATION_JSON));
		HttpEntity<String> entity = new HttpEntity<String>(headers);

		List<LinuxJavaDataSheet> responseRiskScoreDetails = restTemplate()
				.exchange("http://localhost:9090/dataReport", HttpMethod.GET, entity, new ParameterizedTypeReference<List<LinuxJavaDataSheet>>() {}).getBody();
		response.put("dataReportDetail", responseRiskScoreDetails);

		return response;
	}
	
	
	public LinuxJavaDataSheet updateDataReport(int srNo, String comment, Date proposedDate) throws Exception {
		FileInputStream file = new FileInputStream(filepath);

		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheetAt(0);

		//Retrieve the row and check for null
		XSSFRow sheetrow = sheet.getRow(findRow(sheet, srNo));
		if(sheetrow == null){
			throw new Exception("Server not found with SrNo : " + srNo);
		}

		//Update the value of cell
		
		String commentExisting=getStringValue(sheetrow.getCell(27));
		if(commentExisting==null || commentExisting.isEmpty() || commentExisting == "") {
			Cell commentCell = getCellFromRowByNum(sheetrow, COLUMN_COMMENT_NUM);
			commentCell.setCellValue(comment);
		}else if (commentExisting != null) {
			commentExisting=commentExisting +","+ comment;
			Cell commentCell = getCellFromRowByNum(sheetrow, COLUMN_COMMENT_NUM);
			commentCell.setCellValue(commentExisting);
		}
//		commentExisting=commentExisting +","+ comment;
//		Cell commentCell = getCellFromRowByNum(sheetrow, COLUMN_COMMENT_NUM);
//		commentCell.setCellValue(commentExisting);
		Cell proposedDateCell = getCellFromRowByNum(sheetrow, COLUMN_PROPOSED_DATE_NUM);
		proposedDateCell.setCellValue(proposedDate);
		modifyCellTypeAsDate(workbook, proposedDateCell);

		file.close();

		//Update Excel.
		FileOutputStream outFile =new FileOutputStream(new File(filepath));
		workbook.write(outFile);
		outFile.close();

		return populateLinuxJavaDataSheetByRow(sheetrow);
	}
	
	private LinuxJavaDataSheet populateLinuxJavaDataSheetByRow(Row row) {
		LinuxJavaDataSheet dataReportObj = new LinuxJavaDataSheet();
		dataReportObj.setSrNo(getIntegerValue(row.getCell(0)));
		dataReportObj.setPlatform(getStringValue(row.getCell(1)));
		dataReportObj.setServerName(getStringValue(row.getCell(2)));
		dataReportObj.setEnv(getStringValue(row.getCell(3)));
		dataReportObj.setTc(getStringValue(row.getCell(4)));
		dataReportObj.setService(getStringValue(row.getCell(5)));
		dataReportObj.setItsi(getStringValue(row.getCell(6)));
		dataReportObj.setRtbManager(getStringValue(row.getCell(7)));
		dataReportObj.setRtbLead(getStringValue(row.getCell(8)));
		dataReportObj.setIsPrimary(getBooleanValue(row.getCell(9)));
		dataReportObj.setJavaLocation(getStringValue(row.getCell(10)));
		dataReportObj.setJavaClass(getStringValue(row.getCell(11)));
		dataReportObj.setFileVersion(getStringValue(row.getCell(12)));
		dataReportObj.setJavaVersion(getIntegerValue(row.getCell(13)));
		dataReportObj.setJavaType(getStringValue(row.getCell(14)));
		dataReportObj.setPbtCiName(getStringValue(row.getCell(15)));
		dataReportObj.setCommandLastExecuted(getDateValue(row.getCell(16)));
		dataReportObj.setDormancy(getStringValue(row.getCell(17)));
		dataReportObj.setLowCritCount(getIntegerValue(row.getCell(18)));
		dataReportObj.setMedCritCount(getIntegerValue(row.getCell(19)));
		dataReportObj.setHighCritCount(getIntegerValue(row.getCell(20)));
		dataReportObj.setUtilityServer(getStringValue(row.getCell(21)));
		dataReportObj.setUtilityName(getStringValue(row.getCell(22)));
		dataReportObj.setVendor(getStringValue(row.getCell(23)));
		dataReportObj.setEmbeddedType(getStringValue(row.getCell(24)));
		dataReportObj.setJavaClass2(getStringValue(row.getCell(25)));
		dataReportObj.setSuspectedLatestJavaVersion(getStringValue(row.getCell(26)));
		dataReportObj.setComments(getStringValue(row.getCell(27)));
		dataReportObj.setProposedDate(getDateValue(row.getCell(28)));
		return dataReportObj;
	}
	private String getStringValue(Cell cell) {
		return (cell != null && cell.getCellType() == CellType.STRING) ?  cell.getStringCellValue() : "";
	}

	private Date getDateValue(Cell cell) {
		return cell != null ?  cell.getDateCellValue() : null;
	}

	private Integer getIntegerValue(Cell cell) {
		return (cell != null && cell.getCellType() == CellType.NUMERIC) ?  (int)cell.getNumericCellValue() : null;
	}

	private Boolean getBooleanValue(Cell cell) {
		return (cell != null && cell.getCellType() == CellType.BOOLEAN) ?  cell.getBooleanCellValue() : Boolean.FALSE;
	}
	private void modifyCellTypeAsDate(XSSFWorkbook workbook, Cell cell) {
		CellStyle cellStyle = workbook.createCellStyle();
		CreationHelper creationHelper = workbook.getCreationHelper();
		cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat(dateFormat));
		cell.setCellStyle(cellStyle);
	}
	private Cell getCellFromRowByNum(XSSFRow sheetrow, int columnNum) {
		Cell cell = sheetrow.getCell(columnNum);
		if(cell == null){
			cell = sheetrow.createCell(columnNum);
		}
		return cell;
	}
	private static int findRow(XSSFSheet sheet, int cellContent) {
		for (Row row : sheet) {
			for (Cell cell : row) {
				if (cell.getCellType() == CellType.NUMERIC) {
						if (cell.getNumericCellValue()==cellContent) {
						return row.getRowNum();  
					}
				}
			}
		}               
		return -1;
	}

}
/*Cell cell = row.getCell(0);
System.out.println("Cell============="+cell);
if(cell.equals(cellContent)) {
	return row.getRowNum();
}*/
