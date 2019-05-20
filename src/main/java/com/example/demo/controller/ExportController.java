package com.example.demo.controller;

import com.example.demo.entity.Client;
import com.example.demo.service.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.PrintWriter;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;

/**
 * Controlleur pour réaliser les exports.
 */
@Controller
@RequestMapping("/")
public class ExportController {

    @Autowired
    private ClientService clientService;

    @GetMapping("/clients/csv")
    public void clientsCSV(HttpServletRequest request, HttpServletResponse response) throws IOException {

        Integer age;

        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.csv\"");
        PrintWriter writer = response.getWriter();
        List<Client> allClients = clientService.findAllClients();
        LocalDate now = LocalDate.now();
        writer.println("Id;Nom;Prenom;Date de Naissance;Age");

        for(Client client : allClients){
            //NB : Plus le fichier aura de colonnes à remplir et plus ce sera long à coder et susceptible d'avoir des erreurs
            age = now.getYear() - client.getDateNaissance().getYear();
            if (now.getMonthValue() < client.getDateNaissance().getMonthValue()) {
                age -= 1;
            }
            writer.println(client.getId() + ";"
                    + client.getNom().replace(client.getNom(), '"' + client.getNom() + '"') + ";"
                    + client.getPrenom().replace(client.getPrenom(), '"' + client.getPrenom() + '"') + ";"
                    + client.getDateNaissance().format(DateTimeFormatter.ofPattern("dd/MM/yyyy")) + ";"
                    + age);
        }

    }
}
