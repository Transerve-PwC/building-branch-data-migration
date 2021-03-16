package org.egov.building.service;

import java.io.File;

import org.egov.building.model.PropertyResponse;

public interface ReadExcelService {

	public PropertyResponse getDataFromExcel(File file, int sheetIndex);
	public PropertyResponse getDataFromExcelforOwner(File file, int sheetIndex);
	public PropertyResponse getDataFromExcelforDoc(File file, int sheetIndex);
}
