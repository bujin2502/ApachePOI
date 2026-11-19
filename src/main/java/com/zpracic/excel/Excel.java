package com.zpracic.excel;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class Excel {

    public static void main(String[] args) throws IOException {

        // Kreiranje nove Excel radne knjige (.xlsx format)
        XSSFWorkbook radnaKnjiga = new XSSFWorkbook();
        // Dodavanje novog lista (sheet-a) u radnu knjigu
        XSSFSheet list = radnaKnjiga.createSheet("Ocjene studenata");

        // --- STILOVI ĆELIJA ---

        // Stil za zaglavlje tablice (plava pozadina, bijeli bold tekst, centriranje)
        XSSFCellStyle stilZaglavlja = radnaKnjiga.createCellStyle();
        stilZaglavlja.setFillForegroundColor(new XSSFColor(
                new java.awt.Color(46, 117, 182), new DefaultIndexedColorMap())); // Plava pozadina
        stilZaglavlja.setFillPattern(FillPatternType.SOLID_FOREGROUND);           // Puna boja pozadine
        stilZaglavlja.setAlignment(HorizontalAlignment.CENTER);                   // Centriranje teksta
        stilZaglavlja.setBorderBottom(BorderStyle.THIN);                          // Tanka linija ispod
        XSSFFont fontZaglavlja = radnaKnjiga.createFont();
        fontZaglavlja.setBold(true);                                              // Podebljani tekst
        fontZaglavlja.setColor(new XSSFColor(
                new java.awt.Color(255, 255, 255), new DefaultIndexedColorMap())); // Bijela boja teksta
        fontZaglavlja.setFontHeightInPoints((short) 11);                          // Veličina fonta 11pt
        stilZaglavlja.setFont(fontZaglavlja);

        // Stil za parne retke (bijela pozadina, centriran tekst)
        XSSFCellStyle stilParni = radnaKnjiga.createCellStyle();
        stilParni.setAlignment(HorizontalAlignment.CENTER);

        // Stil za neparne retke (svjetloplava pozadina, centriran tekst)
        XSSFCellStyle stilNeparni = radnaKnjiga.createCellStyle();
        stilNeparni.setFillForegroundColor(new XSSFColor(
                new java.awt.Color(222, 234, 246), new DefaultIndexedColorMap())); // Svjetloplava pozadina
        stilNeparni.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        stilNeparni.setAlignment(HorizontalAlignment.CENTER);

        // Stil za redak s formulom prosjeka (podebljano, deblja gornja linija)
        XSSFCellStyle stilZbroj = radnaKnjiga.createCellStyle();
        stilZbroj.setAlignment(HorizontalAlignment.CENTER);
        stilZbroj.setBorderTop(BorderStyle.MEDIUM);  // Deblja linija iznad retka s prosjekom
        XSSFFont fontZbroj = radnaKnjiga.createFont();
        fontZbroj.setBold(true);
        stilZbroj.setFont(fontZbroj);

        // --- ZAGLAVLJE TABLICE ---

        // Nazivi stupaca tablice
        String[] stupci = {"Ime i prezime", "JMBAG", "Kolegij",
                           "Ocjena (broj)", "Ocjena (opis)"};
        // Kreiranje prvog retka (indeks 0) za zaglavlje
        Row redZaglavlja = list.createRow(0);
        for (int i = 0; i < stupci.length; i++) {
            Cell celija = redZaglavlja.createCell(i);
            celija.setCellValue(stupci[i]); // Postavljanje teksta zaglavlja
            celija.setCellStyle(stilZaglavlja); // Primjena stila zaglavlja
        }

        // --- PODACI STUDENATA ---

        // Dvodimenzionalno polje s podacima studenata
        // Svaki redak: ime, JMBAG, kolegij, ocjena (broj), ocjena (opis)
        String[][] podaci = {
            {"Pero Kos",     "0123456789", "Napredne WEB tehnologije", "5", "Izvrstan"},
            {"Franjo Kovač", "0123456790", "Napredne WEB tehnologije", "4", "Vrlo dobar"},
            {"Maja Perić",   "0123456791", "Napredne WEB tehnologije", "5", "Izvrstan"},
            {"Luka Novak",   "0123456792", "Napredne WEB tehnologije", "3", "Dobar"},
            {"Sara Babić",   "0123456793", "Napredne WEB tehnologije", "4", "Vrlo dobar"},
        };

        // Prolaz kroz sve studente i upisivanje podataka u tablicu
        for (int i = 0; i < podaci.length; i++) {
            Row red = list.createRow(i + 1); // +1 jer je redak 0 zaglavlje
            // Naizmjenično se primjenjuje stil za parne i neparne retke
            XSSFCellStyle stil = (i % 2 == 0) ? stilParni : stilNeparni;
            for (int j = 0; j < podaci[i].length; j++) {
                Cell celija = red.createCell(j);
                if (j == 3) {
                    // Stupac "Ocjena (broj)" se upisuje kao cijeli broj, ne kao tekst
                    celija.setCellValue(Integer.parseInt(podaci[i][j]));
                } else {
                    celija.setCellValue(podaci[i][j]); // Ostali stupci su tekst
                }
                celija.setCellStyle(stil);
            }
        }

        // --- REDAK S PROSJEKOM ---

        // Indeks retka odmah ispod zadnjeg studenta
        int zadnjiRed = podaci.length + 1;
        Row redFormule = list.createRow(zadnjiRed);

        // Oznaka "Prosjek ocjena:" u stupcu Kolegij (indeks 2)
        Cell celijaOznaka = redFormule.createCell(2);
        celijaOznaka.setCellValue("Prosjek ocjena:");
        celijaOznaka.setCellStyle(stilZbroj);

        // Excel formula AVERAGE koja računa prosjek ocjena iz stupca D (D2:D6)
        Cell celijaFormula = redFormule.createCell(3);
        celijaFormula.setCellFormula("AVERAGE(D2:D6)");
        celijaFormula.setCellStyle(stilZbroj);

        // --- ZAVRŠNO OBLIKOVANJE ---

        // Automatsko podešavanje širine svakog stupca prema sadržaju
        for (int i = 0; i < stupci.length; i++) {
            list.autoSizeColumn(i);
        }

        // Zamrzavanje prvog retka (zaglavlje ostaje vidljivo pri scrollanju)
        list.createFreezePane(0, 1);

        // --- SPREMANJE DATOTEKE ---

        // Pisanje radne knjige u .xlsx datoteku i automatsko zatvaranje streama
        try (FileOutputStream out = new FileOutputStream("src/main/java/com/zpracic/excel/ocjene.xlsx")) {
            radnaKnjiga.write(out);
        }
        // Oslobađanje resursa radne knjige iz memorije
        radnaKnjiga.close();
        System.out.println("Excel datoteka uspješno kreirana: ocjene.xlsx");
    }
}