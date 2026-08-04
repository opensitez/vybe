*> vybe-test: cobol/xml/xml_generate_with_attributes
*> origin: languages/cobol/tests/cobol/test_xml.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-product.
           05 ws-id    PIC 9(5)  VALUE 1001.
           05 ws-name  PIC X(20) VALUE "Widget".
           05 ws-price PIC 9(5)V99 VALUE 9.99.
       01 ws-xml PIC X(500).
       01 ws-len PIC 9(5).
       PROCEDURE DIVISION.
           XML GENERATE ws-xml FROM ws-product
               COUNT IN ws-len
               WITH ATTRIBUTES
           DISPLAY ws-xml(1:ws-len)
           STOP RUN.

