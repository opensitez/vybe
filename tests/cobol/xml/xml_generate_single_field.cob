*> vybe-test: cobol/xml/xml_generate_single_field
*> origin: languages/cobol/tests/cobol/test_xml.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-title PIC X(30) VALUE "COBOL XML Test".
       01 ws-xml   PIC X(200).
       01 ws-len   PIC 9(5).
       PROCEDURE DIVISION.
           XML GENERATE ws-xml FROM ws-title
               COUNT IN ws-len
           DISPLAY ws-xml(1:ws-len)
           STOP RUN.

