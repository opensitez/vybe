*> vybe-test: cobol/category_xml_json_advanced/test_json_parse_name_mapping
*> origin: languages/cobol/tests/cobol/test_category_xml_json_advanced.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. JSON-NAME-MAP.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 JSON-DOC PIC X(50) VALUE '{"A":1}'.
       01 REC.
          05 A PIC 9.
       PROCEDURE DIVISION.
           JSON PARSE JSON-DOC INTO REC
               NAME A IS "A".
           DISPLAY "JSON MAP PARSED".
    MOVE SPACES TO WS-VYBE-L
    STRING "JSON MAP PARSED" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "JSON MAP PARSED"
        DISPLAY "FAIL: want [JSON MAP PARSED] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

