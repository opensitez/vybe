*> vybe-test: cobol/xml/xml_generate_basic
*> origin: languages/cobol/tests/cobol/test_xml.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-record.
           05 ws-name  PIC X(20) VALUE "Alice".
           05 ws-age   PIC 99    VALUE 30.
       01 ws-xml   PIC X(500).
       01 ws-count PIC 9(5).
       PROCEDURE DIVISION.
           XML GENERATE ws-xml FROM ws-record
               COUNT IN ws-count
           DISPLAY ws-xml(1:ws-count)
           STOP RUN.

