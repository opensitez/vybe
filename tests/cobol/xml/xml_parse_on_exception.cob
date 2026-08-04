*> vybe-test: cobol/xml/xml_parse_on_exception
*> origin: languages/cobol/tests/cobol/test_xml.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-xml  PIC X(50) VALUE "<valid>data</valid>".
       01 ws-err  PIC X     VALUE "N".
       01 ws-ok   PIC X     VALUE "N".
       PROCEDURE DIVISION.
           XML PARSE ws-xml
               PROCESSING PROCEDURE xml-proc
               ON EXCEPTION     MOVE "Y" TO ws-err
               NOT ON EXCEPTION MOVE "Y" TO ws-ok
           END-XML
           DISPLAY ws-ok
           STOP RUN.
       xml-proc SECTION.
           CONTINUE.

