*> vybe-test: cobol/category_intrinsic_math_trig/test_intrinsic_mod_rem
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_math_trig.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. INTRINSIC-MOD.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 RES-MOD PIC 9(2).
       01 RES-REM PIC S9(2).
       PROCEDURE DIVISION.
           COMPUTE RES-MOD = FUNCTION MOD(10, 3).
           COMPUTE RES-REM = FUNCTION REM(-10, 3).
           DISPLAY RES-MOD " " RES-REM.
    MOVE SPACES TO WS-VYBE-L
    STRING RES-MOD DELIMITED SIZE " " DELIMITED SIZE RES-REM DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "01 0q"
        DISPLAY "FAIL: want [01 0q] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

