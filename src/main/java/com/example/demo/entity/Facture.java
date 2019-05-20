package com.example.demo.entity;

import javax.persistence.*;
import java.time.LocalDate;
import java.util.Set;

@Entity
public class Facture {

    @Id
    @GeneratedValue(strategy = GenerationType.AUTO)
    private Long id;

    @ManyToOne
    private Client client;

    @OneToMany(mappedBy = "facture") //Lien avec l'objet facture de LigneFacture
    private Set<LigneFacture> ligneFactures;
    //Le Set<E> empêche les doublons dans sa liste

    @Column
    private LocalDate dateFacture;

    @Column
    private String statut;

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public Client getClient() {
        return client;
    }

    public void setClient(Client client) {
        this.client = client;
    }

    public Set<LigneFacture> getLigneFactures() {
        return ligneFactures;
    }

    public void setLigneFactures(Set<LigneFacture> ligneFactures) {
        this.ligneFactures = ligneFactures;
    }

    public LocalDate getDateFacture() {
        return dateFacture;
    }

    public void setDateFacture(LocalDate dateFacture) {
        this.dateFacture = dateFacture;
    }

    public String getStatut() {
        return statut;
    }

    public void setStatut(String statut) {
        this.statut = statut;
    }

    public Double getTotal() {
        Double total = 0.00;
        for (LigneFacture ligneFacture : ligneFactures) {
            total += ligneFacture.getSousTotal();
        }
        return total;
    }
}
