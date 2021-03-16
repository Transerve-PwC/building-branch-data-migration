package org.egov.building.service;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.stream.Collectors;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;

import org.apache.commons.compress.archivers.dump.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.Styles;
import org.apache.poi.xssf.model.StylesTable;
import org.egov.building.entities.Document;
import org.egov.building.entities.Owner;
import org.egov.building.entities.OwnerDetails;
import org.egov.building.entities.Property;
import org.egov.building.entities.PropertyDetails;
import org.egov.building.model.PropertyResponse;
import org.egov.building.repository.PropertyRepository;
import org.egov.building.service.StreamingSheetContentsHandler.StreamingRowProcessor;
import org.egov.building.util.FileStoreUtils;
import org.egov.tracer.model.CustomException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import lombok.extern.slf4j.Slf4j;

@Slf4j
@Service
public class ReadExcelServiceImpl implements ReadExcelService{

	private static final String SYSTEM = "System";
	private static final String TENANTID = "ch.chandigarh";
	private static final String APPROVE = "APPROVE";
	private static final String ES_PM_APPROVED = "ES_PM_APPROVED";
	private static final String ES_DRAFTED = "ES_DRAFTED";
	private static final String PROPERTY_MASTER = "PROPERTY_MASTER";

	@Autowired
	private PropertyRepository propertyRepository;

	@Autowired
	private FileStoreUtils fileStoreUtils;

	@Value("${file.location}")
	private String fileLocation;

	@Override
	public PropertyResponse getDataFromExcel(File file, int sheetIndex) {
		try {
			OPCPackage opcPackage = OPCPackage.open(file);
			return this.process(opcPackage, sheetIndex);
		} catch (IOException | OpenXML4JException | SAXException e) {
			log.error("Error while parsing Excel", e);
			throw new CustomException("PARSE_ERROR", "Could not parse excel. Error is " + e.getMessage());
		}
	}

	@Override
	public PropertyResponse getDataFromExcelforOwner(File file, int sheetIndex) {
		try {
			OPCPackage opcPackage = OPCPackage.open(file);
			return this.processOwner(opcPackage, sheetIndex);
		} catch (IOException | OpenXML4JException | SAXException e) {
			log.error("Error while parsing Excel", e);
			throw new CustomException("PARSE_ERROR", "Could not parse excel. Error is " + e.getMessage());
		}
	}

	@Override
	public PropertyResponse getDataFromExcelforDoc(File file, int sheetIndex) {
		try {
			OPCPackage opcPackage = OPCPackage.open(file);
			return this.processDoc(opcPackage, sheetIndex);
		} catch (IOException | OpenXML4JException | SAXException e) {
			log.error("Error while parsing Excel", e);
			throw new CustomException("PARSE_ERROR", "Could not parse excel. Error is " + e.getMessage());
		}
	}

	private PropertyResponse process(OPCPackage xlsxPackage, int sheetNo)
			throws IOException, OpenXML4JException, SAXException, CustomException {
		ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(xlsxPackage);
		XSSFReader xssfReader = new XSSFReader(xlsxPackage);
		StylesTable styles = xssfReader.getStylesTable();
		XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
		int index = 0;
		while (iter.hasNext()) {
			try (InputStream stream = iter.next()) {

				if (index == sheetNo) {
					SheetContentsProcessor processor = new SheetContentsProcessor();
					processSheet(styles, strings, new StreamingSheetContentsHandler(processor), stream);
					if (!processor.propertyList.isEmpty()) {
						return saveProperties(processor.propertyList, processor.skippedFileNos);
					} else {
						PropertyResponse propertyResponse = PropertyResponse.builder()
								.skippedFileNos(processor.skippedFileNos)
								.build();
						return propertyResponse;
					}
				}
				index++;
			}
		}
		throw new CustomException("PARSE_ERROR", "Could not process sheet no " + sheetNo);
	}

	private PropertyResponse processOwner(OPCPackage xlsxPackage, int sheetNo)
			throws IOException, OpenXML4JException, SAXException, CustomException {
		ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(xlsxPackage);
		XSSFReader xssfReader = new XSSFReader(xlsxPackage);
		StylesTable styles = xssfReader.getStylesTable();
		XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
		int index = 0;
		while (iter.hasNext()) {
			try (InputStream stream = iter.next()) {

				if (index == sheetNo) {
					SheetContentsProcessorOwner processorOwner = new SheetContentsProcessorOwner();
					processSheet(styles, strings, new StreamingSheetContentsHandler(processorOwner), stream);
					if (!processorOwner.propertyList.isEmpty()) {
						return saveProperties(processorOwner.propertyList, processorOwner.skippedFileNos);
					} else {
						PropertyResponse propertyResponse = PropertyResponse.builder()
								.skippedFileNos(processorOwner.skippedFileNos)
								.build();
						return propertyResponse;
					}
				}
				index++;
			}
		}
		throw new CustomException("PARSE_ERROR", "Could not process sheet no " + sheetNo);
	}

	private PropertyResponse processDoc(OPCPackage xlsxPackage, int sheetNo)
			throws IOException, OpenXML4JException, SAXException, CustomException {
		ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(xlsxPackage);
		XSSFReader xssfReader = new XSSFReader(xlsxPackage);
		StylesTable styles = xssfReader.getStylesTable();
		XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
		int index = 0;
		while (iter.hasNext()) {
			try (InputStream stream = iter.next()) {

				if (index == sheetNo) {
					SheetContentsProcessorDoc processorDoc = new SheetContentsProcessorDoc();
					processSheet(styles, strings, new StreamingSheetContentsHandler(processorDoc), stream);
					if (!processorDoc.propertyList.isEmpty()) {
						PropertyResponse propertyResponse = PropertyResponse.builder()
								.generatedCount(processorDoc.propertyList.size())
								.skippedFileNos(processorDoc.skippedFileNos).build();
						return propertyResponse;
					} else {
						PropertyResponse propertyResponse = PropertyResponse.builder()
								.skippedFileNos(processorDoc.skippedFileNos)
								.build();
						return propertyResponse;
					}
				}
				index++;
			}
		}
		throw new CustomException("PARSE_ERROR", "Could not process sheet no " + sheetNo);
	}

	private class SheetContentsProcessor implements StreamingRowProcessor {

		List<Property> propertyList = new ArrayList<>();
		Set<String> skippedFileNos = new HashSet<>();

		@Override
		public void processRow(Row currentRow) {
			if (currentRow.getRowNum() >= 7) {
				if (currentRow.getCell(1) != null) {
					String firstCell = String
							.valueOf(getValueFromCell(currentRow, 1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK))
							.trim();
					if (isNumeric(firstCell)) {
						//firstCell = firstCell.substring(0, firstCell.length() - 2);
						firstCell = String.valueOf(Double.valueOf(firstCell).intValue());
					}
					Property propertyDb = propertyRepository
								.getPropertyByFileNumber(firstCell);
					
					if (propertyDb == null) {

						int i = 2;
						List<String> excelValues = new ArrayList<>();
						while (i <= 9) {
							excelValues.add(String
									.valueOf(
											getValueFromCell(currentRow, i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK))
									.trim());
							i++;
						}
						if (isNumeric(excelValues.get(7))) {
							PropertyDetails propertyDetails = PropertyDetails.builder()
									.tenantId(TENANTID)
									.houseNumber(excelValues.get(4))
									.mohalla(excelValues.get(5))
									.village(excelValues.get(6))
									.areaSqft(Double.valueOf(excelValues.get(7)).intValue())
									.branchType("BUILDING_BRANCH")
									.build();
							propertyDetails.setCreatedBy(SYSTEM);
							Property property = Property.builder()
									.fileNumber(firstCell)
									.tenantId(TENANTID)
									.propertyMasterOrAllotmentOfSite(PROPERTY_MASTER)
									.action("").state(ES_DRAFTED)
									.category(excelValues.get(0)).subCategory(excelValues.get(1))
									.siteNumber(String.valueOf(Math.round(Float.parseFloat(excelValues.get(2)))))
									.sectorNumber(excelValues.get(3))
									.build();
							property.setCreatedBy(SYSTEM);
							propertyDetails.setProperty(property);
							property.setPropertyDetails(propertyDetails);
							propertyList.add(property);
						} else {
								skippedFileNos.add(firstCell);
								log.error("We are skipping uploading property for file number: " + firstCell
										+ " because of incorrect data.");

						}
					} else {
							skippedFileNos.add(firstCell);
							log.error("We are skipping uploading property for file number: " + firstCell
									+ " as it already exists.");

					}
				}
				}
		}
	}

	private class SheetContentsProcessorOwner implements StreamingRowProcessor {

		List<Property> propertyList = new ArrayList<>();
		Set<String> skippedFileNos = new HashSet<>();

		@Override
		public void processRow(Row currentRow) {
			if (currentRow.getRowNum() >= 7) {
				if (currentRow.getCell(1) != null) {
					String firstCell = String
							.valueOf(getValueFromCell(currentRow, 1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK))
							.trim();
					if (isNumeric(firstCell)) {
						// firstCell = firstCell.substring(0, firstCell.length() - 2);
						firstCell = String.valueOf(Double.valueOf(firstCell).intValue());
					}
					Property propertyDb = propertyRepository.getPropertyByFileNumber(firstCell);
					if (propertyDb != null) {
						int i = 2;
						List<String> excelValues = new ArrayList<>();
						while (i <= 9) {
							excelValues.add(String
									.valueOf(
											getValueFromCell(currentRow, i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK))
									.trim());
							i++;
						}
						if(isNumeric(excelValues.get(5))) {
							OwnerDetails ownerDetails = OwnerDetails.builder()
									.tenantId(TENANTID)
									.ownerName(excelValues.get(0))
									.guardianName(excelValues.get(1))
									.guardianRelation(excelValues.get(2))
									.address(excelValues.get(3))
									.mobileNumber(excelValues.get(4).substring(1, excelValues.get(4).length()-1))
									.isCurrentOwner(Boolean.valueOf(excelValues.get(6)))
									.possesionDate(convertStrDatetoLong(excelValues.get(7)))
									.build();
							ownerDetails.setCreatedBy(SYSTEM);
							Owner owner = Owner.builder()
									.tenantId(TENANTID)
									.share(Double.valueOf(excelValues.get(5)))
									.ownerDetails(ownerDetails)
									.build();
							owner.setCreatedBy(SYSTEM);
							ownerDetails.setOwner(owner);
							owner.setPropertyDetails(propertyDb.getPropertyDetails());
							Set<Owner> owners = new HashSet<>();
							owners.add(owner);
							propertyDb.getPropertyDetails().setOwners(owners);
							propertyRepository.save(propertyDb);
							propertyList.add(propertyDb);
						} else {
							skippedFileNos.add(firstCell);
							log.error("We are skipping uploading owner details for property with file number: " + firstCell
									+ " because of incorrect data.");
						}
					} else {
						skippedFileNos.add(firstCell);
						log.error("We are skipping uploading owner details for property with file number: " + firstCell
								+ " as it does not exists.");
					}
				}
			}
		}
	}

	private void processSheet(Styles styles, SharedStrings strings, SheetContentsHandler sheetHandler,
			InputStream sheetInputStream) throws IOException, SAXException {
		DataFormatter formatter = new DataFormatter();
		InputSource sheetSource = new InputSource(sheetInputStream);
		try {
			SAXParserFactory saxFactory = SAXParserFactory.newInstance();
			saxFactory.setNamespaceAware(false);
			SAXParser saxParser = saxFactory.newSAXParser();
			XMLReader sheetParser = saxParser.getXMLReader();
			ContentHandler handler = new MyXSSFSheetXMLHandler(styles, null, strings, sheetHandler, formatter, false);
			sheetParser.setContentHandler(handler);
			sheetParser.parse(sheetSource);
		} catch (ParserConfigurationException e) {
			throw new RuntimeException("SAX parser appears to be broken - " + e.getMessage());
		}
	}

	private class SheetContentsProcessorDoc implements StreamingRowProcessor {

		Set<String> skippedFileNos = new HashSet<>();
		List<Property> propertyList = new ArrayList<>();

		@Override
		public void processRow(Row currentRow) {
			File folder = new File(fileLocation);
			String[] listOfFiles = folder.list();
			List<String> filesList = Arrays.asList(listOfFiles);

			if (!filesList.isEmpty()) {
			if (currentRow.getRowNum() >= 2) {
				if (currentRow.getCell(0) != null) {
					String firstCell = String
							.valueOf(getValueFromCell(currentRow, 0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK))
							.trim();
					String documentName = String
							.valueOf(getValueFromCell(currentRow, 3, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)).trim();
					if (isNumeric(firstCell)) {
						firstCell = String.valueOf(Double.valueOf(firstCell).intValue());
					}
					if (filesList.contains(documentName)) {
						Property propertyDb = propertyRepository.getPropertyByFileNumber(firstCell);
						if (propertyDb != null) {
							byte[] bytes = null;
							List<HashMap<String, String>> response = null;
							try {
								bytes = Files.readAllBytes(Paths.get(folder + "/" + documentName));
								ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
								outputStream.write(bytes);
								String [] tenantId = propertyDb.getTenantId().split("\\.");
								response = fileStoreUtils.uploadStreamToFileStore(outputStream, tenantId[0],
										documentName);
								outputStream.close();
							} catch (IOException e) {
								log.error("error while converting file into byte output stream");
							}
							String documentType = String
									.valueOf(getValueFromCell(currentRow, 1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)).trim();
							String documentFor = String
									.valueOf(getValueFromCell(currentRow, 4, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)).trim();
							String num = "";
							for (int i = 0; i < documentFor.length(); i++) {
								if (Character.isDigit(documentFor.charAt(i))) {
									num = num + documentFor.charAt(i);
								}
							}
							String docType = "";
							if (documentType.toUpperCase().contains("Sale Deed duly attested by a Notary Public or Executive Magistrate".toUpperCase())) {
								docType = "BB_SALE_DEED";
							} else if (documentType.toUpperCase().contains("Ownership proof".toUpperCase())) {
								docType = "BB_OWNERSHIP_PROOF";
							} else if (documentType.toUpperCase().contains("Report of Revenue".toUpperCase())) {
								docType = "BB_REPORT_OF_REVENUE";
							} else if (documentType.toUpperCase().contains("Photograph of owner".toUpperCase())) {
								docType = "BB_PHOTOGRAPH_OF_OWNER_WITH_BUILDING";
							} else if (documentType.toUpperCase().contains("Building plan".toUpperCase())) {
								docType = "BB_BUILDING_PLAN";
							} else if (documentType.toUpperCase().contains("Affidavit".toUpperCase())) {
								docType = "BB_AFFIDAVIT";
							} else if (documentType.toUpperCase().contains("Indemnity Bond".toUpperCase())) {
								docType = "BB_INDEMNITY_BOND";
							} else if (documentType.toUpperCase().contains("3 Nos. specimen signatures".toUpperCase())) {
								docType = "BB_THREE_NOS_SPECIMEN_SIGNATURE";
							}
							List<Owner> ownerList = propertyDb
									.getPropertyDetails().getOwners().stream().filter(owner -> owner.getOwnerDetails()
											.getIsCurrentOwner().toString().equalsIgnoreCase("true"))
									.collect(Collectors.toList());
							Comparator<Owner> compare = (o1, o2) -> o1.getOwnerDetails().getCreatedTime()
									.compareTo(o2.getOwnerDetails().getCreatedTime());
							Collections.sort(ownerList, compare);
							if(!ownerList.isEmpty()) {
								Document document = Document.builder()
										.tenantId(propertyDb.getTenantId()).active(true).documentType(docType)
										.fileStoreId(response.get(0).get("fileStoreId"))
										.referenceId(ownerList.get(Integer.parseInt(num) - 1).getOwnerDetails().getId())
										.property(propertyDb)
										.build();
								document.setCreatedBy(SYSTEM);
								Set<Document> documents = new HashSet<>();
								documents.add(document);
								propertyDb.setDocuments(documents);
								propertyRepository.save(propertyDb);
								propertyList.add(propertyDb);

							} else {
								skippedFileNos.add(firstCell);
								log.error("We are skipping uploading document for the property having file number: " + firstCell
										+ " as it does not have owner.");
							}
						} else {
							skippedFileNos.add(firstCell);
							log.error("We are skipping uploading property for file number: " + firstCell
									+ " as it does not exists.");
						}
					} else {
						skippedFileNos.add(firstCell);
						log.error("No Document with name "+ documentName + " is present in the document folder");
					}
				}
				}
			} else {
				throw new CustomException("NO_FILES_PRESENT", "No files present in document folder");
			}
		}
	}

	protected Object getValueFromCell(Row row, int cellNo, Row.MissingCellPolicy cellPolicy) {
		Cell cell1 = row.getCell(cellNo, cellPolicy);
		Object objValue = "";
		switch (cell1.getCellType()) {
		case BLANK:
			objValue = "";
			break;
		case STRING:
			objValue = cell1.getStringCellValue();
			break;
		case NUMERIC:
			try {
				if (DateUtil.isCellDateFormatted(cell1)) {
					objValue = cell1.getDateCellValue().getTime();
				} else {
					throw new InvalidFormatException();
				}
			} catch (Exception ex1) {
				try {
					objValue = cell1.getNumericCellValue();
				} catch (Exception ex2) {
					objValue = 0.0;
				}
			}

			break;
		case FORMULA:
			objValue = cell1.getNumericCellValue();
			break;

		default:
			objValue = "";
		}
		return objValue;
	}

	protected long convertStrDatetoLong(String dateStr) {
		try {
			SimpleDateFormat f = new SimpleDateFormat("dd/MM/yyyy");
			Date d = f.parse(dateStr);
			return d.getTime();
		} catch (Exception e) {
			log.error("Date parsing issue occur :" + e.getMessage());
		}
		return 0;
	}

	private PropertyResponse saveProperties(List<Property> properties, Set<String> skippedFileNos) {
		properties.forEach(property -> {
			propertyRepository.save(property);
		});
		PropertyResponse propertyResponse = PropertyResponse.builder().generatedCount(properties.size())
				.skippedFileNos(skippedFileNos).build();
		return propertyResponse;
	}

	private Boolean isNumeric(String value) {
		if (value != null && !value.matches("[1-9][0-9]*(\\.[0])?")) {
			return false;
		}
		return true;
	}
}
