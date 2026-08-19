*> vybe-test: cobol/category_json_xml/test_json_parse_basic
*> origin: languages/cobol/tests/cobol/test_category_json_xml.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. JSON-PARSE.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 JSON-DOC PIC X(20) VALUE '{"A":"HI"}'.
       01 REC.
          05 A PIC X(2).
       PROCEDURE DIVISION.
           JSON PARSE JSON-DOC INTO REC.
           DISPLAY A.
    MOVE SPACES TO WS-VYBE-L
    STRING A DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = ""
        DISPLAY "FAIL: want [] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

