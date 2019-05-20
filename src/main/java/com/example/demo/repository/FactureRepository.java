package com.example.demo.repository;

import com.example.demo.entity.Facture;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import java.util.List;

@Repository
public interface FactureRepository extends JpaRepository<Facture, Long> {

    //@Query("Select f FROM Facture f WHERE f.client.id = ?1")
    List<Facture> findByClientId(Long clientId);
}
