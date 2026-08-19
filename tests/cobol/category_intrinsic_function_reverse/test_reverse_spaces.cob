*> vybe-test: cobol/category_intrinsic_function_reverse/test_reverse_spaces
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_function_reverse.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V PIC X(5) VALUE 'A B C'. PROCEDURE DIVISION. DISPLAY FUNCTION REVERSE(V).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION REVERSE(V) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "C B A"
        DISPLAY "FAIL: want [C B A] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

