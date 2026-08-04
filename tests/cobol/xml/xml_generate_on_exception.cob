*> vybe-test: cobol/xml/xml_generate_on_exception
*> origin: languages/cobol/tests/cobol/test_xml.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-rec.
           05 ws-val PIC X(5) VALUE "test".
       01 ws-xml    PIC X(50).
       01 ws-len    PIC 9(5).
       01 ws-err    PIC X VALUE "N".
       PROCEDURE DIVISION.
           XML GENERATE ws-xml FROM ws-rec
               COUNT IN ws-len
               ON EXCEPTION     MOVE "Y" TO ws-err
               NOT ON EXCEPTION MOVE "N" TO ws-err
           END-XML
           DISPLAY ws-err
           STOP RUN.

