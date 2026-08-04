*> vybe-test: cobol/category_unstring_overflow/test_unstr_delim_in
*> origin: languages/cobol/tests/cobol/test_category_unstring_overflow.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 S PIC X(5) VALUE 'A*B-C'. 01 R1 PIC X. 01 D1 PIC X. PROCEDURE DIVISION. UNSTRING S DELIMITED BY '*' OR '-' INTO R1 DELIMITER IN D1. DISPLAY D1.
    MOVE SPACES TO WS-VYBE-L
    STRING D1 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "*"
        DISPLAY "FAIL: want [*] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

