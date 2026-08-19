*> vybe-test: cobol/category_intrinsics/test_intrinsic_numval_c
*> origin: languages/cobol/tests/cobol/test_category_intrinsics.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. NUMVAL-C-TEST.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 STR PIC X(15) VALUE "  $1,234.56CR".
       01 RES PIC S9(4)V99.
       PROCEDURE DIVISION.
           COMPUTE RES = FUNCTION NUMVAL-C(STR).
           DISPLAY RES.
    MOVE SPACES TO WS-VYBE-L
    STRING RES DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "12345v"
        DISPLAY "FAIL: want [12345v] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

