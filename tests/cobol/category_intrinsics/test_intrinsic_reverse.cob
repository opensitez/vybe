*> vybe-test: cobol/category_intrinsics/test_intrinsic_reverse
*> origin: languages/cobol/tests/cobol/test_category_intrinsics.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. REVERSE-TEST.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 STR PIC X(5) VALUE "COBOL".
       01 RES PIC X(5).
       PROCEDURE DIVISION.
           MOVE FUNCTION REVERSE(STR) TO RES.
           DISPLAY RES.
    MOVE SPACES TO WS-VYBE-L
    STRING RES DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "LOBOC"
        DISPLAY "FAIL: want [LOBOC] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

