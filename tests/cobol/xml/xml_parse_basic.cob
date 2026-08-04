*> vybe-test: cobol/xml/xml_parse_basic
*> origin: languages/cobol/tests/cobol/test_xml.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-xml-doc PIC X(200)
           VALUE "<person><name>Alice</name><age>30</age></person>".
       01 ws-event-count PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           XML PARSE ws-xml-doc
               PROCESSING PROCEDURE xml-handler
           DISPLAY ws-event-count
           STOP RUN.
       xml-handler SECTION.
           ADD 1 TO ws-event-count.

