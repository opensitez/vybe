*> vybe-test: cobol/category_perform/test_perform_out_of_line_basic
*> origin: languages/cobol/tests/cobol/test_category_perform.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. PERFORM-BASIC.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 COUNTER PIC 9 VALUE 0.
       PROCEDURE DIVISION.
           PERFORM PARA-A.
           DISPLAY "END".
    MOVE SPACES TO WS-VYBE-L
    STRING "END" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "END"
        DISPLAY "FAIL: want [END] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.
       PARA-A.
           DISPLAY "PARA-A".
    MOVE SPACES TO WS-VYBE-L
    STRING "PARA-A" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "PARA-A"
        DISPLAY "FAIL: want [PARA-A] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           EXIT.

