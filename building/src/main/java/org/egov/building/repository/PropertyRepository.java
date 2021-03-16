package org.egov.building.repository;

import org.egov.building.entities.Property;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface PropertyRepository extends JpaRepository<Property, String>{

	public Property getPropertyByFileNumber(String fileNumber);
}
