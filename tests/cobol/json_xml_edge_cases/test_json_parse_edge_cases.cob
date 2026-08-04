*> vybe-test: cobol/json_xml_edge_cases/test_json_parse_edge_cases
*> origin: languages/cobol/tests/cobol/test_json_xml_edge_cases.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. JSON-EDGE-CASES.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 JSON-DOC PIC X(200).
       01 JSON-STATUS PIC 9(4).
       01 PARSED-DATA.
          05 USER-AGE PIC 9(3).
          05 IS-ACTIVE PIC X(5).
          05 NULL-FIELD PIC X(10).
       PROCEDURE DIVISION.
           MOVE '{"USER-AGE": 42, "IS-ACTIVE": true, "NULL-FIELD": null}' TO JSON-DOC.
           JSON PARSE JSON-DOC INTO PARSED-DATA
                ON EXCEPTION
                   DISPLAY "EXCEPTION:" JSON-STATUS
                NOT ON EXCEPTION
                   DISPLAY USER-AGE " " IS-ACTIVE " " NULL-FIELD
           END-JSON.
           STOP RUN.

