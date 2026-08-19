*> vybe-test: cobol/category_intrinsics/test_intrinsic_numval
*> origin: languages/cobol/tests/cobol/test_category_intrinsics.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. NUMVAL-TEST.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 STR PIC X(10) VALUE "  -123.45 ".
       01 RES PIC S9(4)V99.
       PROCEDURE DIVISION.
           COMPUTE RES = FUNCTION NUMVAL(STR).
           DISPLAY RES.
    MOVE SPACES TO WS-VYBE-L
    STRING RES DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "01234u"
        DISPLAY "FAIL: want [01234u] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

