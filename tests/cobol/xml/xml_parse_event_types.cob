*> vybe-test: cobol/xml/xml_parse_event_types
*> origin: languages/cobol/tests/cobol/test_xml.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-xml PIC X(200)
           VALUE "<book><title>COBOL Guide</title><pages>400</pages></book>".
       01 ws-start-count   PIC 99 VALUE 0.
       01 ws-end-count     PIC 99 VALUE 0.
       01 ws-content-count PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           XML PARSE ws-xml
               PROCESSING PROCEDURE parse-events
           DISPLAY ws-start-count
           DISPLAY ws-end-count
           DISPLAY ws-content-count
           STOP RUN.
       parse-events SECTION.
           EVALUATE XML-CODE
               WHEN "START-OF-ELEMENT"
                   ADD 1 TO ws-start-count
               WHEN "END-OF-ELEMENT"
                   ADD 1 TO ws-end-count
               WHEN "CONTENT-CHARACTERS"
                   ADD 1 TO ws-content-count
               WHEN OTHER
                   CONTINUE
           END-EVALUATE.

