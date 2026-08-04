*> vybe-test: cobol/category_data_division_external_clause/test_external_parse_6
*> origin: languages/cobol/tests/cobol/test_category_data_division_external_clause.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 EX-GRP. 05 X PIC X VALUE 'Z'. 05 Y PIC X VALUE 'Q' IS EXTERNAL. PROCEDURE DIVISION. DISPLAY X.
    MOVE SPACES TO WS-VYBE-L
    STRING X DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "Z"
        DISPLAY "FAIL: want [Z] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. DISPLAY Y.
    MOVE SPACES TO WS-VYBE-L
    STRING Y DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "Q"
        DISPLAY "FAIL: want [Q] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

