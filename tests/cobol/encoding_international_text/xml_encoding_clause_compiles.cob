*> vybe-test: cobol/encoding_international_text/xml_encoding_clause_compiles
*> origin: languages/cobol/tests/cobol/test_encoding_international_text.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 R PIC X(10).
01 X PIC X(100).
01 L PIC 9(5).
PROCEDURE DIVISION.
    XML GENERATE X FROM R COUNT IN L ENCODING 1208.
    STOP RUN.

