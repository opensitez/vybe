*> vybe-test: cobol/xml/xml_generate_numeric_field
*> origin: languages/cobol/tests/cobol/test_xml.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-amount PIC 9(7)V99 VALUE 12345.67.
       01 ws-xml    PIC X(200).
       01 ws-len    PIC 9(5).
       PROCEDURE DIVISION.
           XML GENERATE ws-xml FROM ws-amount
               COUNT IN ws-len
           DISPLAY ws-xml(1:ws-len)
           STOP RUN.

