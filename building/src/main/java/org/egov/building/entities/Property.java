package org.egov.building.entities;

import java.util.HashSet;
import java.util.Set;

import javax.persistence.CascadeType;
import javax.persistence.Column;
import javax.persistence.EmbeddedId;
import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.Id;
import javax.persistence.OneToMany;
import javax.persistence.OneToOne;
import javax.persistence.Table;

import org.hibernate.annotations.GenericGenerator;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import lombok.ToString;

@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
@ToString
@Builder
@Entity
@Table(name = "cs_ep_property_v1")
public class Property extends AuditDetails {

	@Id
	@GeneratedValue(generator = "UUID")
	@GenericGenerator(
			name = "UUID",
			strategy = "org.hibernate.id.UUIDGenerator"
	)
	@Column(name = "id")
	private String id;

	@Column(name = "tenantid")
	private String tenantId;

	@Column(name = "file_number")
	private String fileNumber;

	/**
	 * One of the categories from `data/ch/EstateServices/Categories.json`
	 * CAT.RESIDENTIAL, CAT.COMMERCIAL, CAT.INDUSTRIAL, CAT.INSTITUTIONAL,
	 * CAT.GOVPROPERTY, CAT.RELIGIOUS, CAT.HOSPITAL,
	 *
	 */
	@Column(name = "category")
	private String category;

	@Column(name = "sub_category")
	private String subCategory;

	@Column(name = "site_number")
	private String siteNumber;

	@Column(name = "sector_number")
	private String sectorNumber;

	@Column(name = "state")
	private String state;

	@Column(name = "action")
	private String action;

	@OneToOne(
            cascade =  CascadeType.ALL,
            mappedBy = "property")
	private PropertyDetails propertyDetails;

	@OneToMany(
			cascade = CascadeType.ALL,
			mappedBy = "property"
			)
	private Set<Document> documents = new HashSet<>();

	@Column(name = "property_master_or_allotment_of_site")
	private String propertyMasterOrAllotmentOfSite;

}
