package com.example.demo.service;

import com.example.demo.entity.Facture;
import com.example.demo.repository.FactureRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.util.List;

@Service
@Transactional
public class FactureService {

    @Autowired
    private FactureRepository factureRepository;

    public List<Facture> findAllFactures() {
        return factureRepository.findAll();
    }

    public List<Facture> findFacturesClient(Long clientId) {
        return factureRepository.findByClientId(clientId);
    }
}
