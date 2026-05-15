package com.zpracic.word;

import java.io.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import org.apache.poi.xwpf.usermodel.*;

public class CirkularnoPismo {

    public static void main(String[] args) throws IOException {
        String csvDatoteka = "src/main/java/com/zpracic/word/studenti.csv";
        String predlozakDatoteka = "src/main/java/com/zpracic/word/predlozak.docx";
        String datum = LocalDate.now()
                .format(DateTimeFormatter.ofPattern("dd.MM.yyyy."));

        try (BufferedReader br = new BufferedReader(
                new FileReader(csvDatoteka))) {
            String linija;
            boolean prviRed = true;
            while ((linija = br.readLine()) != null) {
                if (prviRed) {
                    prviRed = false;
                    continue;
                }
                String[] podaci = linija.split(",");
                String ime     = podaci[0];
                String prezime = podaci[1];
                String jmbag   = podaci[2];
                String kolegij = podaci[3];
                String ocjena  = podaci[4];

                try (FileInputStream fis = new FileInputStream(predlozakDatoteka);
                     XWPFDocument dokument = new XWPFDocument(fis)) {

                    for (XWPFParagraph paragraf : dokument.getParagraphs()) {
                        zamijeniUParagrafu(paragraf, ime, prezime, 
                                jmbag, kolegij, ocjena, datum);
                    }

                    String izlazDatoteka = "src/main/java/com/zpracic/word/obavijest_"
                            + ime + "_" + prezime + ".docx";
                    try (FileOutputStream out =
                            new FileOutputStream(izlazDatoteka)) {
                        dokument.write(out);
                    }
                    System.out.println("Kreiran: " + izlazDatoteka);
                }
            }
        }
        System.out.println("Cirkularno pismo uspješno generirano.");
    }

    private static void zamijeniUParagrafu(XWPFParagraph paragraf,
            String ime, String prezime, String jmbag,
            String kolegij, String ocjena, String datum) {

        StringBuilder cijeli = new StringBuilder();
        for (XWPFRun run : paragraf.getRuns()) {
            if (run.getText(0) != null) {
                cijeli.append(run.getText(0));
            }
        }

        String tekst = cijeli.toString();
        tekst = tekst.replace("{{IME}}", ime);
        tekst = tekst.replace("{{PREZIME}}", prezime);
        tekst = tekst.replace("{{JMBAG}}", jmbag);
        tekst = tekst.replace("{{KOLEGIJ}}", kolegij);
        tekst = tekst.replace("{{OCJENA}}", ocjena);
        tekst = tekst.replace("{{DATUM}}", datum);

        if (!paragraf.getRuns().isEmpty()) {
            paragraf.getRuns().get(0).setText(tekst, 0);
            for (int i = 1; i < paragraf.getRuns().size(); i++) {
                paragraf.getRuns().get(i).setText("", 0);
            }
        }
    }
}