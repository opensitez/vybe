*> vybe-test: cobol/paragraph_section_flow/paragraph_called_conditionally_true
*> origin: languages/cobol/tests/cobol/test_paragraph_section_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9 VALUE 5.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF N > 0
        PERFORM SAY-POS
    END-IF.
    STOP RUN.
SAY-POS.
    DISPLAY "POSITIVE".
    MOVE SPACES TO WS-VYBE-L
    STRING "POSITIVE" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "POSITIVE"
        DISPLAY "FAIL: want [POSITIVE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

