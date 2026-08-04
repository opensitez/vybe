*> vybe-test: cobol/xml/xml_generate_special_chars
*> origin: languages/cobol/tests/cobol/test_xml.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-data.
           05 ws-note PIC X(30) VALUE "Price < 100 & > 50".
       01 ws-xml PIC X(500).
       01 ws-len PIC 9(5).
       PROCEDURE DIVISION.
           XML GENERATE ws-xml FROM ws-data
               COUNT IN ws-len
           DISPLAY ws-xml(1:ws-len)
           STOP RUN.

