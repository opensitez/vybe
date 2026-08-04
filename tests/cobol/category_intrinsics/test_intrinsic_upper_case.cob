*> vybe-test: cobol/category_intrinsics/test_intrinsic_upper_case
*> origin: languages/cobol/tests/cobol/test_category_intrinsics.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. UPPER-CASE-TEST.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 STR PIC X(10) VALUE "hello".
       01 RES PIC X(10).
       PROCEDURE DIVISION.
           MOVE FUNCTION UPPER-CASE(STR) TO RES.
           DISPLAY RES.
    MOVE SPACES TO WS-VYBE-L
    STRING RES DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "HELLO     "
        DISPLAY "FAIL: want [HELLO     ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

