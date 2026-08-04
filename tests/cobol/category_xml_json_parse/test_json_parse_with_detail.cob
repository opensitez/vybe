*> vybe-test: cobol/category_xml_json_parse/test_json_parse_with_detail
*> origin: languages/cobol/tests/cobol/test_category_xml_json_parse.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 DOC PIC X(20) VALUE '{"A":1}'. 01 R. 05 A PIC 9. PROCEDURE DIVISION. JSON PARSE DOC INTO R WITH DETAIL. DISPLAY 'OK'.
    MOVE SPACES TO WS-VYBE-L
    STRING 'OK' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "OK"
        DISPLAY "FAIL: want [OK] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

