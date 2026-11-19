package com.zpracic.word;

import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.*;

public class Word {
    public static void main(String[] args) throws IOException {

        // Kreiranje novog Word dokumenta (.docx format)
        XWPFDocument dokument = new XWPFDocument();

        // --- ZAGLAVLJE DOKUMENTA ---

        // Kreiranje zadanog (default) zaglavlja koje se pojavljuje na svakoj stranici
        XWPFHeader zaglavlje = dokument.createHeader(HeaderFooterType.DEFAULT);
        XWPFParagraph zaglavljeParagraf = zaglavlje.createParagraph();
        zaglavljeParagraf.setAlignment(ParagraphAlignment.CENTER); // Centriranje teksta zaglavlja
        XWPFRun zaglavljeRun = zaglavljeParagraf.createRun();
        zaglavljeRun.setText("Fakultet organizacije i informatike - Varaždin");
        zaglavljeRun.setFontSize(10);
        zaglavljeRun.setColor("888888"); // Siva boja teksta

        // --- PODNOŽJE DOKUMENTA ---

        // Kreiranje zadanog (default) podnožja koje se pojavljuje na svakoj stranici
        XWPFFooter podnozje = dokument.createFooter(HeaderFooterType.DEFAULT);
        XWPFParagraph podnozjeParagraf = podnozje.createParagraph();
        podnozjeParagraf.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun podnozjeRun = podnozjeParagraf.createRun();
        podnozjeRun.setText("Seminarski rad — Apache POI");
        podnozjeRun.setFontSize(10);
        podnozjeRun.setColor("888888"); // Siva boja teksta

        // --- NASLOV DOKUMENTA ---

        // Kreiranje paragrafa za glavni naslov
        XWPFParagraph naslov = dokument.createParagraph();
        naslov.setAlignment(ParagraphAlignment.CENTER);
        naslov.setSpacingAfter(200); // Razmak ispod naslova (u twipsima, 200 = ~3.5mm)
        XWPFRun naslovRun = naslov.createRun();
        naslovRun.setText("Popis studenata i ocjena");
        naslovRun.setBold(true);
        naslovRun.setFontSize(20);
        naslovRun.setFontFamily("Calibri");

        // --- PODNASLOV ---

        // Kreiranje paragrafa za podnaslov s akademskom godinom
        XWPFParagraph podnaslov = dokument.createParagraph();
        podnaslov.setAlignment(ParagraphAlignment.CENTER);
        podnaslov.setSpacingAfter(400); // Veći razmak ispod podnaslova prije teksta
        XWPFRun podnaslovRun = podnaslov.createRun();
        podnaslovRun.setText("Akademska godina 2025./2026.");
        podnaslovRun.setItalic(true);   // Kurziv
        podnaslovRun.setFontSize(12);
        podnaslovRun.setColor("444444"); // Tamnija siva boja

        // --- UVODNI TEKST ---

        // Kratki opisni paragraf iznad tablice
        XWPFParagraph uvod = dokument.createParagraph();
        uvod.setSpacingAfter(200);
        XWPFRun uvodRun = uvod.createRun();
        uvodRun.setText("U nastavku se nalazi popis studenata s postignutim ocjenama "
                + "na kolegiju Napredne WEB tehnologije i servisi. "
                + "Tablica prikazuje ime i prezime studenta, JMBAG te završnu ocjenu.");
        uvodRun.setFontSize(11);
        uvodRun.setFontFamily("Calibri");

        // --- PODACI ZA TABLICU ---

        // Dvodimenzionalno polje s podacima studenata
        // Svaki redak: ime i prezime, JMBAG, ocjena
        String[][] podaci = {
                { "Emanuel Pračić", "0123456789", "Izvrstan (5)" },
                { "Pero Kos",       "0123456789", "Izvrstan (5)" },
                { "Franjo Kovač",   "0123456790", "Vrlo dobar (4)" },
                { "Maja Perić",     "0123456791", "Izvrstan (5)" },
                { "Luka Novak",     "0123456792", "Dobar (3)" },
                { "Sara Babić",     "0123456793", "Vrlo dobar (4)" }
        };

        // --- KREIRANJE TABLICE ---

        // Tablica ima (broj studenata + 1) redaka — +1 zbog retka zaglavlja
        // i onoliko stupaca koliko ima podataka po studentu
        XWPFTable tablica = dokument.createTable((podaci.length + 1), podaci[0].length);
        tablica.setWidth("100%"); // Tablica zauzima cijelu širinu stranice

        // --- ZAGLAVLJE TABLICE ---

        String[] zaglavljaStupaca = { "Ime i prezime", "JMBAG", "Ocjena" };
        XWPFTableRow zaglavljeRed = tablica.getRow(0); // Dohvat prvog retka (index 0)
        for (int i = 0; i < zaglavljaStupaca.length; i++) {
            XWPFTableCell celija = zaglavljeRed.getCell(i);
            celija.setColor("2E75B6"); // Plava pozadina ćelije zaglavlja
            XWPFParagraph p = celija.getParagraphs().get(0);
            p.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun r = p.createRun();
            r.setText(zaglavljaStupaca[i]);
            r.setBold(true);
            r.setColor("FFFFFF"); // Bijeli tekst
            r.setFontSize(11);
        }

        // --- REDCI S PODACIMA STUDENATA ---

        for (int i = 0; i < podaci.length; i++) {
            XWPFTableRow red = tablica.getRow(i + 1); // +1 preskačemo redak zaglavlja
            // Naizmjenična bijela i svjetloplava boja retka (zebra efekt)
            String boja = (i % 2 == 0) ? "FFFFFF" : "DEEAF1";
            for (int j = 0; j < podaci[i].length; j++) {
                XWPFTableCell celija = red.getCell(j);
                celija.setColor(boja); // Primjena boje pozadine
                XWPFParagraph p = celija.getParagraphs().get(0);
                p.setAlignment(ParagraphAlignment.CENTER);
                XWPFRun r = p.createRun();
                r.setText(podaci[i][j]);
                r.setFontSize(11);
            }
        }

        // --- ZAVRŠNA NAPOMENA ---

        // Paragraf ispod tablice s napomenom o načinu generiranja dokumenta
        XWPFParagraph kraj = dokument.createParagraph();
        kraj.setSpacingBefore(400); // Razmak iznad završnog teksta
        XWPFRun krajRun = kraj.createRun();
        krajRun.setText("Dokument generiran programski korištenjem Apache POI biblioteke.");
        krajRun.setItalic(true);
        krajRun.setFontSize(10);
        krajRun.setColor("888888"); // Siva boja — vizualno manje istaknut tekst

        // --- SPREMANJE DOKUMENTA ---

        // Pisanje dokumenta u .docx datoteku i automatsko zatvaranje streama
        try (FileOutputStream out = new FileOutputStream("src/main/java/com/zpracic/word/studenti.docx")) {
            dokument.write(out);
        }
        // Oslobađanje resursa dokumenta iz memorije
        dokument.close();
        System.out.println("Word dokument uspješno kreiran: studenti.docx");
    }
}