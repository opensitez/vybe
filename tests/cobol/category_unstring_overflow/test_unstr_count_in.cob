*> vybe-test: cobol/category_unstring_overflow/test_unstr_count_in
*> origin: languages/cobol/tests/cobol/test_category_unstring_overflow.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 S PIC X(5) VALUE 'AB*CD'. 01 R1 PIC X(2). 01 C1 PIC 9. PROCEDURE DIVISION. UNSTRING S DELIMITED BY '*' INTO R1 COUNT IN C1. DISPLAY C1.
    MOVE SPACES TO WS-VYBE-L
    STRING C1 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "2"
        DISPLAY "FAIL: want [2] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

