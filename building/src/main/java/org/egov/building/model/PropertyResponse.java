package org.egov.building.model;

import java.util.HashSet;
import java.util.Set;

import javax.validation.Valid;

import com.fasterxml.jackson.annotation.JsonProperty;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Builder.Default;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor
@Builder
public class PropertyResponse {

	@JsonProperty("Inserted")
	@Valid
	@Default
	private int generatedCount = 0;

	@JsonProperty("Skipped Properties File Numbers")
	@Valid
	private Set<String> skippedFileNos = new HashSet<>();
}
