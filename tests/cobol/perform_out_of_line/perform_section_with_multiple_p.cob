*> vybe-test: cobol/perform_out_of_line/perform_section_with_multiple_paras
*> origin: languages/cobol/tests/cobol/test_perform_out_of_line.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    PERFORM WORK-SEC.
    STOP RUN.
WORK-SEC SECTION.
    DISPLAY "A".
    MOVE SPACES TO WS-VYBE-L
    STRING "A" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "A"
        DISPLAY "FAIL: want [A] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    DISPLAY "B".
    MOVE SPACES TO WS-VYBE-L
    STRING "B" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "B"
        DISPLAY "FAIL: want [B] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

