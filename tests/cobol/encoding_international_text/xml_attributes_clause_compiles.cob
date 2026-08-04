*> vybe-test: cobol/encoding_international_text/xml_attributes_clause_compiles
*> origin: languages/cobol/tests/cobol/test_encoding_international_text.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 R.
   05 A PIC X(5).
01 X PIC X(100).
01 L PIC 9(5).
PROCEDURE DIVISION.
    XML GENERATE X FROM R COUNT IN L WITH ATTRIBUTES.
    STOP RUN.

