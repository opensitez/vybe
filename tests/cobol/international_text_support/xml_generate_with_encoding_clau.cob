*> vybe-test: cobol/international_text_support/xml_generate_with_encoding_clause_compiles
*> origin: languages/cobol/tests/cobol/test_international_text_support.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-REC.
   05 WS-NAME PIC X(10) VALUE "ALICE".
01 WS-XML PIC X(200).
01 WS-LEN PIC 9(5).
PROCEDURE DIVISION.
    XML GENERATE WS-XML FROM WS-REC COUNT IN WS-LEN ENCODING 1208.
    STOP RUN.

