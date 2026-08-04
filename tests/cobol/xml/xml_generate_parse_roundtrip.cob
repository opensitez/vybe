*> vybe-test: cobol/xml/xml_generate_parse_roundtrip
*> origin: languages/cobol/tests/cobol/test_xml.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-original.
           05 ws-id   PIC 9(5) VALUE 42.
           05 ws-name PIC X(10) VALUE "Alice".
       01 ws-xml      PIC X(500).
       01 ws-xml-len  PIC 9(5).
       01 ws-events   PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           XML GENERATE ws-xml FROM ws-original
               COUNT IN ws-xml-len
           XML PARSE ws-xml(1:ws-xml-len)
               PROCESSING PROCEDURE count-events
           DISPLAY ws-events
           STOP RUN.
       count-events SECTION.
           ADD 1 TO ws-events.

