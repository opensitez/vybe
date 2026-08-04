*> vybe-test: cobol/perform_out_of_line/perform_thru_three_paragraphs
*> origin: languages/cobol/tests/cobol/test_perform_out_of_line.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    PERFORM A THRU C.
    STOP RUN.
A.
    DISPLAY "A".
    MOVE SPACES TO WS-VYBE-L
    STRING "A" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "A"
        DISPLAY "FAIL: want [A] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
B.
    DISPLAY "B".
    MOVE SPACES TO WS-VYBE-L
    STRING "B" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "B"
        DISPLAY "FAIL: want [B] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
C.
    DISPLAY "C".
    MOVE SPACES TO WS-VYBE-L
    STRING "C" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "C"
        DISPLAY "FAIL: want [C] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

