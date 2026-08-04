*> vybe-test: cobol/category_intrinsic_advanced/test_intrinsic_date_of_integer
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_advanced.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. INTRINSIC-DOI.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 RES PIC 9(8).
       PROCEDURE DIVISION.
           COMPUTE RES = FUNCTION DATE-OF-INTEGER(1).
           DISPLAY RES.
    MOVE SPACES TO WS-VYBE-L
    STRING RES DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "16010101"
        DISPLAY "FAIL: want [16010101] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

