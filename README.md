# Apache POI Demo

Seminarski rad – primjeri korištenja biblioteke [Apache POI](https://poi.apache.org/) za programsko kreiranje Microsoft Office dokumenata (Excel i Word) u Javi.

## Sadržaj projekta

| Klasa | Opis | Izlaz |
|-------|------|-------|
| `com.zpracic.excel.Excel` | Kreira Excel tablicu s ocjenama studenata: stilizirano zaglavlje, zebra retci, formula `AVERAGE` za prosjek, automatska širina stupaca i zamrznut prvi redak. | `src/main/java/com/zpracic/excel/ocjene.xlsx` |
| `com.zpracic.word.Word` | Kreira Word dokument sa zaglavljem, podnožjem, naslovom, podnaslovom i formatiranom tablicom studenata. | `src/main/java/com/zpracic/word/studenti.docx` |
| `com.zpracic.word.CirkularnoPismo` | Cirkularno pismo (mail merge): čita podatke iz `studenti.csv` i za svakog studenta popunjava predložak `predlozak.docx` zamjenom oznaka `{{IME}}`, `{{PREZIME}}`, `{{JMBAG}}`, `{{KOLEGIJ}}`, `{{OCJENA}}` i `{{DATUM}}`. | `src/main/java/com/zpracic/word/obavijest_<Ime>_<Prezime>.docx` |

```
.
├── pom.xml
└── src/main/java/com/zpracic
    ├── excel
    │   └── Excel.java
    └── word
        ├── CirkularnoPismo.java
        ├── Word.java
        ├── predlozak.docx   # predložak za cirkularno pismo
        └── studenti.csv     # ulazni podaci za cirkularno pismo
```

## Preduvjeti

- **JDK 25** (projekt se prevodi za Java 25, vidi `pom.xml`)
- **Maven 3.9+**

Korištene ovisnosti (preuzima ih Maven automatski):

- `org.apache.poi:poi` 5.5.1
- `org.apache.poi:poi-ooxml` 5.5.1

## Pokretanje

Naredbe treba pokretati iz **korijenskog direktorija projekta** jer programi koriste relativne putanje do ulaznih i izlaznih datoteka.

```bash
# Prevođenje
mvn compile

# Excel primjer
mvn exec:java -Dexec.mainClass=com.zpracic.excel.Excel

# Word primjer
mvn exec:java -Dexec.mainClass=com.zpracic.word.Word

# Cirkularno pismo
mvn exec:java -Dexec.mainClass=com.zpracic.word.CirkularnoPismo
```

### Pokretanje iz IDE-a

Projekt je standardni Maven projekt pa se može uvesti u bilo koji IDE:

- **Eclipse:** *File → Import → Maven → Existing Maven Projects* i odabrati direktorij projekta. Eclipse će sam generirati svoje konfiguracijske datoteke (`.project`, `.classpath`, `.settings/`), koje se namjerno ne nalaze u repozitoriju.
- **IntelliJ IDEA:** *File → Open* i odabrati `pom.xml`.

Zatim pokrenuti željenu klasu kao *Java Application* (radni direktorij mora biti korijen projekta).

## Generirane datoteke

Izlazne `.xlsx` i `.docx` datoteke nastaju pokretanjem programa i nisu dio repozitorija (navedene su u `.gitignore`).
