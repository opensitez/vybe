*> vybe-test: cobol/perform_out_of_line/perform_paragraph_exits_correctly
*> origin: languages/cobol/tests/cobol/test_perform_out_of_line.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 FLAG PIC X VALUE "N".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    PERFORM CHECK-FLAG.
    DISPLAY "AFTER".
    MOVE SPACES TO WS-VYBE-L
    STRING "AFTER" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "AFTER"
        DISPLAY "FAIL: want [AFTER] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.
CHECK-FLAG.
    MOVE "Y" TO FLAG.
    STOP RUN.

