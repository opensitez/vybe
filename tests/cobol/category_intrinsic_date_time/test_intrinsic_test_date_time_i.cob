*> vybe-test: cobol/category_intrinsic_date_time/test_intrinsic_test_date_time_invalid_format
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_date_time.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. INTRINSIC-DTINVALID.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 RES PIC 9.
       PROCEDURE DIVISION.
           COMPUTE RES = FUNCTION TEST-DATE-TIME("NOTDATE", "%Y%m%d").
           DISPLAY RES.
    MOVE SPACES TO WS-VYBE-L
    STRING RES DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "0"
        DISPLAY "FAIL: want [0] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

