package org.egov.building.util;

import java.io.ByteArrayOutputStream;
import java.io.UnsupportedEncodingException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.stereotype.Repository;
import org.springframework.util.LinkedMultiValueMap;
import org.springframework.util.MultiValueMap;
import org.springframework.web.client.RestTemplate;

import lombok.extern.slf4j.Slf4j;

@Slf4j
@Repository
public class FileStoreUtils {

	@Value("${egov.filestore.host}${egov.file.url.path}")
	private String fileStoreUrl;

	private RestTemplate restTemplate;

	public FileStoreUtils(RestTemplate restTemplate) {
		super();
		this.restTemplate = restTemplate;
	}

	@SuppressWarnings("unchecked")
	public List<HashMap<String, String>> uploadStreamToFileStore(ByteArrayOutputStream outputStream, String tenantId,
			String fileName) throws UnsupportedEncodingException {
		StringBuilder uri = new StringBuilder(fileStoreUrl.substring(0, fileStoreUrl.length() - 4));

		HttpHeaders headers = new HttpHeaders();
		headers.setContentType(MediaType.MULTIPART_FORM_DATA);
		MultiValueMap<String, Object> body = new LinkedMultiValueMap<>();
		HttpEntity<MultiValueMap<String, Object>> requestEntity = new HttpEntity<>(body, headers);

		body.add("filename", fileName);

		ByteArrayResource contentsAsResource = new ByteArrayResource(outputStream.toByteArray()) {
			@Override
			public String getFilename() {
				return fileName; // Filename has to be returned in order to be able to post.
			}
		};
		body.add("file", contentsAsResource);

		uri.append("?tenantId=" + tenantId + "&module=" + "RentedProperties");
		try {
			Map<String, Map<String, String>> response = (Map<String, Map<String, String>>) restTemplate
					.postForObject(uri.toString(), requestEntity, HashMap.class);

			List<HashMap<String, String>> result = (List<HashMap<String, String>>) response.get("files");
			return result;
		} catch (Exception e) {
			log.error("Exception while fetching file store id: ", e);
		}
		return null;
	}

}
