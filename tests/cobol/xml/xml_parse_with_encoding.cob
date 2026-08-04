*> vybe-test: cobol/xml/xml_parse_with_encoding
*> origin: languages/cobol/tests/cobol/test_xml.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-xml PIC X(100)
           VALUE "<root><item>value</item></root>".
       01 ws-count PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           XML PARSE ws-xml
               PROCESSING PROCEDURE handle-xml
               ENCODING 1208
           DISPLAY ws-count
           STOP RUN.
       handle-xml SECTION.
           ADD 1 TO ws-count.

