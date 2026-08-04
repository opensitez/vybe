*> vybe-test: cobol/xml/xml_generate_with_xml_declaration
*> origin: languages/cobol/tests/cobol/test_xml.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-item.
           05 ws-name PIC X(10) VALUE "widget".
           05 ws-qty  PIC 99    VALUE 5.
       01 ws-xml PIC X(500).
       01 ws-len PIC 9(5).
       PROCEDURE DIVISION.
           XML GENERATE ws-xml FROM ws-item
               COUNT IN ws-len
               WITH XML-DECLARATION
           DISPLAY ws-xml(1:ws-len)
           STOP RUN.

