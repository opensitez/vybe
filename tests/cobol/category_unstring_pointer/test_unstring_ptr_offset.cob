*> vybe-test: cobol/category_unstring_pointer/test_unstring_ptr_offset
*> origin: languages/cobol/tests/cobol/test_category_unstring_pointer.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 S1 PIC X(5) VALUE 'A,B,C'. 01 R1 PIC X. 01 P PIC 9 VALUE 3. PROCEDURE DIVISION. UNSTRING S1 DELIMITED BY ',' INTO R1 WITH POINTER P. DISPLAY R1.
    MOVE SPACES TO WS-VYBE-L
    STRING R1 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "B"
        DISPLAY "FAIL: want [B] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

