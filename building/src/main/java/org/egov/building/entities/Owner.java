package org.egov.building.entities;

import javax.persistence.CascadeType;
import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.FetchType;
import javax.persistence.GeneratedValue;
import javax.persistence.Id;
import javax.persistence.JoinColumn;
import javax.persistence.ManyToOne;
import javax.persistence.OneToOne;
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
@Table(name = "cs_ep_owner_v1")
public class Owner extends AuditDetails {

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

	@ManyToOne(fetch = FetchType.LAZY)
	@JoinColumn(name = "property_details_id")
	private PropertyDetails propertyDetails;

	@Column(name = "share")
	private double share;

	@OneToOne(
			cascade = CascadeType.ALL,
			mappedBy = "owner")
	private OwnerDetails ownerDetails;
	
}
