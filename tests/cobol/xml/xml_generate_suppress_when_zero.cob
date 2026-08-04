*> vybe-test: cobol/xml/xml_generate_suppress_when_zero
*> origin: languages/cobol/tests/cobol/test_xml.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-order.
           05 ws-order-id  PIC 9(5) VALUE 500.
           05 ws-qty       PIC 99   VALUE 0.
           05 ws-amount    PIC 9(7)V99 VALUE 0.
       01 ws-xml PIC X(500).
       01 ws-len PIC 9(5).
       PROCEDURE DIVISION.
           XML GENERATE ws-xml FROM ws-order
               COUNT IN ws-len
               SUPPRESS WHEN ZERO
           DISPLAY ws-xml(1:ws-len)
           STOP RUN.

