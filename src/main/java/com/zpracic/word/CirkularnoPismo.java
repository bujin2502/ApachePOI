package com.zpracic.word;

import java.io.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import org.apache.poi.xwpf.usermodel.*;

public class CirkularnoPismo {
    public static void main(String[] args) throws IOException {

        // Putanja do CSV datoteke s podacima studenata
        String csvDatoteka = "src/main/java/com/zpracic/word/studenti.csv";
        // Putanja do Word predloška s oznakama za zamjenu (npr. {{IME}}, {{OCJENA}}...)
        String predlozakDatoteka = "src/main/java/com/zpracic/word/predlozak.docx";

        // Dohvat trenutnog datuma i formatiranje u oblik dd.MM.yyyy. (npr. 23.05.2026.)
        String datum = LocalDate.now()
                .format(DateTimeFormatter.ofPattern("dd.MM.yyyy."));

        // Otvaranje CSV datoteke za čitanje redak po redak
        try (BufferedReader br = new BufferedReader(new FileReader(csvDatoteka))) {
            String linija;
            boolean prviRed = true; // Zastavica za preskakanje zaglavlja CSV-a

            while ((linija = br.readLine()) != null) {

                // Preskače se prvi redak jer sadrži nazive stupaca, ne podatke
                if (prviRed) {
                    prviRed = false;
                    continue;
                }

                // Razdvajanje retka po zarezu u pojedinačne podatke
                String[] podaci = linija.split(",");
                String ime     = podaci[0];
                String prezime = podaci[1];
                String jmbag   = podaci[2];
                String kolegij = podaci[3];
                String ocjena  = podaci[4];

                // Za svakog studenta otvara se svježi primjerak predloška
                // (try-with-resources automatski zatvara stream nakon upotrebe)
                try (FileInputStream fis = new FileInputStream(predlozakDatoteka);
                     XWPFDocument dokument = new XWPFDocument(fis)) {

                    // Prolaz kroz sve paragrafe dokumenta i zamjena oznaka stvarnim podacima
                    for (XWPFParagraph paragraf : dokument.getParagraphs()) {
                        zamijeniUParagrafu(paragraf, ime, prezime,
                                jmbag, kolegij, ocjena, datum);
                    }

                    // Generiranje jedinstvenog naziva izlazne datoteke za svakog studenta
                    String izlazDatoteka = "src/main/java/com/zpracic/word/obavijest_"
                            + ime + "_" + prezime + ".docx";

                    // Spremanje popunjenog dokumenta u novu .docx datoteku
                    try (FileOutputStream out = new FileOutputStream(izlazDatoteka)) {
                        dokument.write(out);
                    }

                    System.out.println("Kreiran: " + izlazDatoteka);
                }
            }
        }

        System.out.println("Cirkularno pismo uspješno generirano.");
    }

    /**
     * Zamjenjuje sve oznake (placeholdere) u jednom paragrafu stvarnim podacima studenta.
     *
     * Problem: Apache POI može razbiti jedan vidljivi tekst u više Run objekata
     * (npr. "{{IME}}" može biti split na "{{", "IME", "}}"), što otežava zamjenu.
     * Rješenje: spajaju se svi Runovi u jedan string, radi se zamjena, pa se vraća
     * cijeli tekst u prvi Run, a ostale Runove se prazne.
     */
    private static void zamijeniUParagrafu(XWPFParagraph paragraf,
            String ime, String prezime, String jmbag,
            String kolegij, String ocjena, String datum) {

        // Spajanje teksta iz svih Runova paragrafa u jedan string
        StringBuilder cijeli = new StringBuilder();
        for (XWPFRun run : paragraf.getRuns()) {
            if (run.getText(0) != null) {
                cijeli.append(run.getText(0));
            }
        }

        // Zamjena svih oznaka (placeholdera) odgovarajućim podacima studenta
        String tekst = cijeli.toString();
        tekst = tekst.replace("{{IME}}",     ime);
        tekst = tekst.replace("{{PREZIME}}", prezime);
        tekst = tekst.replace("{{JMBAG}}",   jmbag);
        tekst = tekst.replace("{{KOLEGIJ}}", kolegij);
        tekst = tekst.replace("{{OCJENA}}",  ocjena);
        tekst = tekst.replace("{{DATUM}}",   datum);

        // Upisivanje zamijenjenog teksta natrag u paragraf:
        // - cijeli tekst ide u prvi Run (indeks 0)
        // - svi ostali Runovi se prazne kako bi se izbjeglo dupliciranje sadržaja
        if (!paragraf.getRuns().isEmpty()) {
            paragraf.getRuns().get(0).setText(tekst, 0);
            for (int i = 1; i < paragraf.getRuns().size(); i++) {
                paragraf.getRuns().get(i).setText("", 0);
            }
        }
    }
}