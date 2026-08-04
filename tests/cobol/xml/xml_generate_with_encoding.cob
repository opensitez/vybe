*> vybe-test: cobol/xml/xml_generate_with_encoding
*> origin: languages/cobol/tests/cobol/test_xml.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-data.
           05 ws-id    PIC 9(5) VALUE 42.
           05 ws-label PIC X(10) VALUE "item".
       01 ws-xml  PIC X(500).
       01 ws-len  PIC 9(5).
       PROCEDURE DIVISION.
           XML GENERATE ws-xml FROM ws-data
               COUNT IN ws-len
               ENCODING 1208
           DISPLAY ws-xml(1:ws-len)
           STOP RUN.

