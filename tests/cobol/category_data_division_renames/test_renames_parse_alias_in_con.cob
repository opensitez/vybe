*> vybe-test: cobol/category_data_division_renames/test_renames_parse_alias_in_condition_true
*> origin: languages/cobol/tests/cobol/test_category_data_division_renames.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 R. 05 A PIC X VALUE 'Q'. 05 B PIC X VALUE 'R'. 05 C PIC X VALUE 'S'. 66 ALIAS RENAMES A THRU B. PROCEDURE DIVISION. IF ALIAS = 'QR' DISPLAY 'MATCH' ELSE DISPLAY 'MISS' END-IF STOP RUN.
    MOVE SPACES TO WS-VYBE-L
    STRING 'MATCH' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "MATCH"
        DISPLAY "FAIL: want [MATCH] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

