package org.egov.building.entities;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.FetchType;
import javax.persistence.GeneratedValue;
import javax.persistence.Id;
import javax.persistence.JoinColumn;
import javax.persistence.ManyToOne;
import javax.persistence.Table;

import org.hibernate.annotations.GenericGenerator;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.EqualsAndHashCode;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
@Builder
@EqualsAndHashCode
@Entity
@Table(name = "cs_ep_documents_v1")
public class Document extends AuditDetails{

	@Id
	@GeneratedValue(generator = "UUID")
	@GenericGenerator(
			name = "UUID",
			strategy = "org.hibernate.id.UUIDGenerator"
	)
	@Column(name = "id")
	private String id;

	@Column(name = "reference_id")
	private String referenceId;

	@Column(name = "tenantid")
	private String tenantId;

	@Column(name = "is_active")
	private Boolean active;

	@Column(name = "document_type")
	private String documentType;

	@Column(name = "file_store_id")
	private String fileStoreId;

	@ManyToOne(fetch = FetchType.LAZY)
	@JoinColumn(name = "property_id")
	private Property property;

}
