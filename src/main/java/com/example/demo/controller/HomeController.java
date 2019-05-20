package com.example.demo.controller;

import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.service.ClientService;
import com.example.demo.service.FactureService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.servlet.ModelAndView;

import java.util.List;

/**
 * Controller principale pour affichage des clients / factures sur la page d'acceuil.
 */
@Controller
public class HomeController {

    @Autowired
    private ClientService clientService;

    @Autowired
    private FactureService factureService;

    @GetMapping("/")
    public ModelAndView home() {
        ModelAndView modelAndView = new ModelAndView("home");

        List<Client> clientList = clientService.findAllClients();
        modelAndView.addObject("clients", clientList);

        List<Facture> factureList = factureService.findAllFactures();
        modelAndView.addObject("factures", factureList);

        return modelAndView;
    }
}
