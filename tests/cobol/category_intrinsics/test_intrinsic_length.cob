*> vybe-test: cobol/category_intrinsics/test_intrinsic_length
*> origin: languages/cobol/tests/cobol/test_category_intrinsics.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. LENGTH-TEST.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 STR PIC X(15) VALUE "TEST".
       01 LEN PIC 9(2).
       PROCEDURE DIVISION.
           COMPUTE LEN = FUNCTION LENGTH(STR).
           DISPLAY LEN.
    MOVE SPACES TO WS-VYBE-L
    STRING LEN DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "15"
        DISPLAY "FAIL: want [15] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

