*> vybe-test: cobol/category_data_division_blank_when_zero/test_bwz_zero_after_move
*> origin: languages/cobol/tests/cobol/test_category_data_division_blank_when_zero.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 X PIC 9(4) VALUE 1234 BLANK WHEN ZERO. PROCEDURE DIVISION. MOVE 0 TO X DISPLAY '[' X ']'.
    MOVE SPACES TO WS-VYBE-L
    STRING '[' DELIMITED SIZE X DELIMITED SIZE ']' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "[    ]"
        DISPLAY "FAIL: want [[    ]] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

