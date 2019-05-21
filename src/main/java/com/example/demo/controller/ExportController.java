package com.example.demo.controller;

import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.entity.LigneFacture;
import com.example.demo.service.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Optional;

/**
 * Controlleur pour réaliser les exports.
 */
@Controller
@RequestMapping("/")
public class ExportController {

    @Autowired
    private ClientService clientService;

    @Autowired
    private FactureService factureService;

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

    @GetMapping("/clients/xlsx")
    public void clientsXLSX(HttpServletRequest request, HttpServletResponse response) throws IOException {

        response.setHeader("Content-Disposition", "attachment; filename=\"clients.xlsx\"");

        // Creating a Workbook from an Excel file (.xls or .xlsx)
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Clients");

        //Création de la ligne des en-têtes
        Row headerRowEntete = sheet.createRow(0);
        headerRowEntete.createCell(0).setCellValue("Nom");
        headerRowEntete.createCell(1).setCellValue("Prénom");
        headerRowEntete.createCell(2).setCellValue("Date de naissance");

        Integer i = 1;
        List<Client> allClients = clientService.findAllClients();
        for (Client client : allClients) {
            Row headerRow = sheet.createRow(i);
            headerRow.createCell(0).setCellValue(client.getNom());
            headerRow.createCell(1).setCellValue(client.getPrenom());
            headerRow.createCell(2).setCellValue(client.getDateNaissance().format(DateTimeFormatter.ofPattern("dd/MM/yyyy")));
            i += 1;
        }
        workbook.write(response.getOutputStream());
        workbook.close();
    }

    @GetMapping("/clients/{idClient}/factures/xlsx")
    public void factureClientXLSX(@PathVariable Long idClient, HttpServletRequest request, HttpServletResponse response) throws IOException {

        response.setHeader("Content-Disposition", "attachment; filename=\"factureClient.xlsx\"");

        // Creating a Workbook from an Excel file (.xls or .xlsx)
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Factures");

        //Création de la ligne des en-têtes
        Row headerRowEntete = sheet.createRow(0);
        headerRowEntete.createCell(0).setCellValue("N° facture");
        headerRowEntete.createCell(1).setCellValue("Prix total");

        Integer i = 1;
        List<Facture> allFactures = factureService.findFacturesClient(idClient);
        for (Facture facture : allFactures) {
            Row headerRow = sheet.createRow(i);
            headerRow.createCell(0).setCellValue(facture.getId());
            headerRow.createCell(1).setCellValue(facture.getTotal());
            i += 1;
        }
        workbook.write(response.getOutputStream());
        workbook.close();
    }

    @GetMapping("/factures/xlsx")
    public void facturesXLSX(HttpServletRequest request, HttpServletResponse response) throws IOException {

        //Créer un onglet par facture en créant avant chaque facture de préciser la fiche client (infos en ligne)
        // Entête (si facture) : nom article, qté, prix unitaire, prix de la ligne
        //  gras, rouge avec bordure pour toute la ligne total en fusionnant avec les cellules vides de la ligne pour avoir le mot 'TOTAL' aligné à côté du total
        //Si client, mettre l'entête en colonne : nom, prenom, date de naissance
        response.setHeader("Content-Disposition", "attachment; filename=\"factures.xlsx\"");

        // Creating a Workbook from an Excel file (.xls or .xlsx)
        Workbook workbook = new XSSFWorkbook();

        List<Client> allClients = clientService.findAllClients();
        for(Client client : allClients) {
            createClientSheet(client, workbook);
            List<Facture> allFactures = factureService.findFacturesClient(client.getId());
            for (Facture facture : allFactures) {
                createFactureSheet(facture, workbook);
            }
        }
        workbook.write(response.getOutputStream());
        workbook.close();
    }

    private void createClientSheet(Client client, Workbook workbook) {
        //Création d'une feuille par chaque client associé aux feuilles de facture qui suivent dans le fichier
        Sheet sheet = workbook.createSheet("Client " + client.getId());

        //Création de la ligne des en-têtes
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Nom");
        headerRow.createCell(1).setCellValue(client.getNom());
        sheet.autoSizeColumn(0);

        headerRow = sheet.createRow(1);
        headerRow.createCell(0).setCellValue("Prénom");
        headerRow.createCell(1).setCellValue(client.getPrenom());
        sheet.autoSizeColumn(1);

        headerRow = sheet.createRow(2);
        headerRow.createCell(0).setCellValue("Date de naissance");
        headerRow.createCell(1).setCellValue(client.getDateNaissance().format(DateTimeFormatter.ofPattern("dd/MM/yyyy")));
        sheet.autoSizeColumn(2);

    }

    private Workbook createFactureSheet(Facture facture, Workbook workbook) {

        //Création d'une feuille par facture
        Sheet sheet = workbook.createSheet("Facture " + facture.getId());

        //Création de la ligne des en-têtes
        Integer i = 0;
        Row headerRowEntete = sheet.createRow(i);
        headerRowEntete.createCell(0).setCellValue("Nom article");
        headerRowEntete.createCell(1).setCellValue("Quantité");
        headerRowEntete.createCell(2).setCellValue("Prix unitaire");
        headerRowEntete.createCell(3).setCellValue("Prix de la ligne");
        sheet.autoSizeColumn(i);

        //Création d'une ligne par article
        for(LigneFacture ligneFacture : facture.getLigneFactures()){
            i += 1;
            Row headerRow = sheet.createRow(i);
            headerRow.createCell(0).setCellValue(ligneFacture.getArticle().getLibelle());
            headerRow.createCell(1).setCellValue(ligneFacture.getQuantite());
            headerRow.createCell(2).setCellValue(ligneFacture.getArticle().getPrix());
            headerRow.createCell(3).setCellValue(ligneFacture.getSousTotal());
            sheet.autoSizeColumn(i);
        }
        //Ligne de fin de la facture : TOTAL
        Row headerRowFooter = sheet.createRow(i+1);
        headerRowFooter.createCell(2).setCellValue("Total");
        headerRowFooter.createCell(3).setCellValue(facture.getTotal());
        sheet.autoSizeColumn(i+1);

        // Create a Font for styling header cells
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setColor(IndexedColors.RED.getIndex());

        return workbook;
    }
}
