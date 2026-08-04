*> vybe-test: cobol/json_xml_edge_cases/test_json_parse_with_name_clause
*> origin: languages/cobol/tests/cobol/test_json_xml_edge_cases.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. JSON-NAME-EDGE.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 JSON-DOC PIC X(200) VALUE '{"A": 42, "B": true}'.
       01 PARSED-DATA.
          05 USER-AGE PIC 9(3).
          05 IS-ACTIVE PIC X(5).
       PROCEDURE DIVISION.
           JSON PARSE JSON-DOC INTO PARSED-DATA
              NAME USER-AGE IS "A"
              NAME IS-ACTIVE IS "B".
           DISPLAY "JSON NAME EDGE".
    MOVE SPACES TO WS-VYBE-L
    STRING "JSON NAME EDGE" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "JSON NAME EDGE"
        DISPLAY "FAIL: want [JSON NAME EDGE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

