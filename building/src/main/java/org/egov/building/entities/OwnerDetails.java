package org.egov.building.entities;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.FetchType;
import javax.persistence.GeneratedValue;
import javax.persistence.Id;
import javax.persistence.JoinColumn;
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
@ToString
@AllArgsConstructor
@NoArgsConstructor
@Builder
@Table
@Entity(name = "cs_ep_owner_details_v1")
public class OwnerDetails extends AuditDetails {

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

	@OneToOne(fetch = FetchType.LAZY)
	@JoinColumn(name = "owner_id")
	private Owner owner;

	@Column(name = "owner_name")
	private String ownerName;

	@Column(name = "guardian_name")
	private String guardianName;

	@Column(name = "guardian_relation")
	private String guardianRelation;

	@Column(name = "mobile_number")
	private String mobileNumber;

	@Column(name = "possesion_date")
	private Long possesionDate;

	@Column(name = "address")
	private String address;

	@Column(name = "is_current_owner")
	private Boolean isCurrentOwner;

}
