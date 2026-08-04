*> vybe-test: cobol/category_data_division_renames/test_renames_parse_alias_compare_false
*> origin: languages/cobol/tests/cobol/test_category_data_division_renames.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 R. 05 A PIC X VALUE '9'. 05 B PIC X VALUE '9'. 05 C PIC X VALUE '9'. 66 ALIAS RENAMES B THRU C. PROCEDURE DIVISION. IF ALIAS = '99' DISPLAY 'OK' ELSE DISPLAY 'BAD' END-IF STOP RUN.
    MOVE SPACES TO WS-VYBE-L
    STRING 'OK' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "OK"
        DISPLAY "FAIL: want [OK] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

