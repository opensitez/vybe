*> vybe-test: cobol/category_unstring_overflow/test_unstr_pointer_update
*> origin: languages/cobol/tests/cobol/test_category_unstring_overflow.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 S PIC X(5) VALUE 'A*B*C'. 01 R1 PIC X. 01 P PIC 9 VALUE 1. PROCEDURE DIVISION. UNSTRING S DELIMITED BY '*' INTO R1 WITH POINTER P. DISPLAY P.
    MOVE SPACES TO WS-VYBE-L
    STRING P DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "3"
        DISPLAY "FAIL: want [3] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

