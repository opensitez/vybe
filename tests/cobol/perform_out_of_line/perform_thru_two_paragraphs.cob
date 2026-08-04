*> vybe-test: cobol/perform_out_of_line/perform_thru_two_paragraphs
*> origin: languages/cobol/tests/cobol/test_perform_out_of_line.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    PERFORM P1 THRU P2.
    STOP RUN.
P1.
    DISPLAY "P1".
    MOVE SPACES TO WS-VYBE-L
    STRING "P1" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "P1"
        DISPLAY "FAIL: want [P1] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
P2.
    DISPLAY "P2".
    MOVE SPACES TO WS-VYBE-L
    STRING "P2" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "P2"
        DISPLAY "FAIL: want [P2] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

