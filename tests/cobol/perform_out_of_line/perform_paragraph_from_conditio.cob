*> vybe-test: cobol/perform_out_of_line/perform_paragraph_from_conditional_branch
*> origin: languages/cobol/tests/cobol/test_perform_out_of_line.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9 VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    PERFORM SET-X.
    IF X = 99
        DISPLAY "OK"
    ELSE
        DISPLAY "FAIL"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "OK" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "OK"
        DISPLAY "FAIL: want [OK] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.
SET-X.
    MOVE 99 TO X.
    STOP RUN.

