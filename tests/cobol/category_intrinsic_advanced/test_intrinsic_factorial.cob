*> vybe-test: cobol/category_intrinsic_advanced/test_intrinsic_factorial
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_advanced.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. INTRINSIC-FACT.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 RES PIC 9(4).
       PROCEDURE DIVISION.
           COMPUTE RES = FUNCTION FACTORIAL(5).
           DISPLAY RES.
    MOVE SPACES TO WS-VYBE-L
    STRING RES DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "0120"
        DISPLAY "FAIL: want [0120] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

