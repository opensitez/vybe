*> vybe-test: cobol/category_intrinsic_function_length/test_length_group
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_function_length.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 G. 05 A PIC X(2). 05 B PIC X(3). PROCEDURE DIVISION. DISPLAY FUNCTION LENGTH(G).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION LENGTH(G) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "5"
        DISPLAY "FAIL: want [5] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

