package org.egov.building.controller;

import java.io.File;

import org.apache.commons.lang3.StringUtils;
import org.egov.building.model.PropertyResponse;
import org.egov.building.service.ReadExcelService;
import org.egov.tracer.model.CustomException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import lombok.extern.slf4j.Slf4j;

@Slf4j
@Controller
@RequestMapping("/v1/excel")
public class ReadExcelController {

	private ReadExcelService readExcelService;

	@Autowired
	public ReadExcelController(ReadExcelService readExcelService) {
		this.readExcelService = readExcelService;
	}

	@Value("${file.path}")
	private String filePath;

	@PostMapping("/read")
	public ResponseEntity<?> readExcel() {
		try {
			log.info("Start controller method readExcel() Request:" + filePath);
			if (StringUtils.isBlank(filePath)) {
				throw new Exception("Cannot find property file that is uploaded");
			}
			File tempFile = new File(filePath);
			if(!tempFile.exists()) {
				throw new CustomException("FILE_NOT_FOUND", "File not found in resource folder");
			}
			PropertyResponse propertyResponse = this.readExcelService.getDataFromExcel(tempFile, 1);
			log.info("End controller method readExcel property inserted:" + propertyResponse.getGeneratedCount());

			return new ResponseEntity<>(propertyResponse, HttpStatus.OK);
		} catch (Exception e) {	
			log.error("Error occurred during readExcel():" + e.getMessage(), e);
			throw new CustomException("FILE_TEMPLATE_NOT_VALID", "Invalid template uploaded. Please upload a valid property excel file.");
		}

	}

	@PostMapping("/read_owner")
	public ResponseEntity<?> readExcelOwner() {
		try {
			log.info("Start controller method readExcelOwner() Request:" + filePath);
			if (StringUtils.isBlank(filePath)) {
				throw new Exception("Cannot find property file that is uploaded");
			}
			File tempFile = new File(filePath);
			if(!tempFile.exists()) {
				throw new CustomException("FILE_NOT_FOUND", "File not found in resource folder");
			}
			PropertyResponse propertyResponse = this.readExcelService.getDataFromExcelforOwner(tempFile, 2);
			log.info("End controller method readExcelOwner property inserted:" + propertyResponse.getGeneratedCount());

			return new ResponseEntity<>(propertyResponse, HttpStatus.OK);
		} catch (Exception e) {	
			log.error("Error occurred during readExcel():" + e.getMessage(), e);
			throw new CustomException("FILE_TEMPLATE_NOT_VALID", "Invalid template uploaded. Please upload a valid property excel file.");
		}

	}

	@PostMapping("/read_doc")
	public ResponseEntity<?> readExcelDoc() {
		try {
			log.info("Start controller method readExcelDoc() Request:" + filePath);
			if (StringUtils.isBlank(filePath)) {
				throw new Exception("Cannot find property file that is uploaded");
			}
			File tempFile = new File(filePath);
			if(!tempFile.exists()) {
				throw new CustomException("FILE_NOT_FOUND", "File not found in resource folder");
			}
			PropertyResponse propertyResponse = this.readExcelService.getDataFromExcelforDoc(tempFile, 3);
			log.info("End controller method readExcelDoc property inserted:" + propertyResponse.getGeneratedCount());

			return new ResponseEntity<>(propertyResponse, HttpStatus.OK);
		} catch (Exception e) {
			log.error("Error occurred during readExcel():" + e.getMessage(), e);
			throw new CustomException("FILE_TEMPLATE_NOT_VALID", "Invalid template uploaded. Please upload a valid property excel file.");
		}

	}

}
