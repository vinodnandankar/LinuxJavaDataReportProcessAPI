package com.example.demo.controller;

import java.util.Date;
import java.util.Map;

import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PutMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import com.example.demo.entity.LinuxJavaDataSheet;
import com.example.demo.service.DataReportPApiService;

@RestController
public class DataReportPApiController {
	
	public final DataReportPApiService dataReportPApiService;
	
	public DataReportPApiController(DataReportPApiService dataReportPApiService) {
		this.dataReportPApiService = dataReportPApiService;
	}

	@GetMapping("/getReportDetail")
	public Map<String, Object> getDataReport() {
		return dataReportPApiService.getDataReport();
	}
	
	@PutMapping("/dataReport/update")
	public ResponseEntity<LinuxJavaDataSheet> updateDataReport(@RequestParam int srNo, @RequestParam String comment, @RequestParam Date proposedDate) {
		try {
			return new ResponseEntity<>(dataReportPApiService.updateDataReport(srNo, comment, proposedDate), HttpStatus.ACCEPTED);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return new ResponseEntity<>(HttpStatus.INTERNAL_SERVER_ERROR);
	}}
