*> vybe-test: cobol/xml/xml_generate_suppress_when_spaces
*> origin: languages/cobol/tests/cobol/test_xml.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-contact.
           05 ws-name  PIC X(20) VALUE "Alice".
           05 ws-phone PIC X(15) VALUE SPACES.
           05 ws-email PIC X(30) VALUE SPACES.
       01 ws-xml PIC X(500).
       01 ws-len PIC 9(5).
       PROCEDURE DIVISION.
           XML GENERATE ws-xml FROM ws-contact
               COUNT IN ws-len
               SUPPRESS WHEN SPACES
           DISPLAY ws-xml(1:ws-len)
           STOP RUN.

