*> vybe-test: cobol/category_data_division_renames/test_renames_parse_alias_with_trailing_spaces
*> origin: languages/cobol/tests/cobol/test_category_data_division_renames.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 R. 05 A PIC X VALUE 'A'. 05 B PIC X VALUE ' '. 05 C PIC X VALUE 'C'. 66 ALIAS RENAMES A THRU B. PROCEDURE DIVISION. MOVE 'T ' TO ALIAS DISPLAY ALIAS.
    MOVE SPACES TO WS-VYBE-L
    STRING ALIAS DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "T "
        DISPLAY "FAIL: want [T ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

