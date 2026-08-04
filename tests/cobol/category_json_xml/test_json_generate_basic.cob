*> vybe-test: cobol/category_json_xml/test_json_generate_basic
*> origin: languages/cobol/tests/cobol/test_category_json_xml.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. JSON-GEN.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 REC.
          05 FLD-A PIC X(5) VALUE "HELLO".
          05 FLD-B PIC 9(3) VALUE 123.
       01 JSON-DOC PIC X(50).
       PROCEDURE DIVISION.
           JSON GENERATE JSON-DOC FROM REC.
           DISPLAY "JSON GEN PARSED".
    MOVE SPACES TO WS-VYBE-L
    STRING "JSON GEN PARSED" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "JSON GEN PARSED"
        DISPLAY "FAIL: want [JSON GEN PARSED] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

