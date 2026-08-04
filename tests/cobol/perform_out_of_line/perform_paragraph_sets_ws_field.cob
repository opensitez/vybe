*> vybe-test: cobol/perform_out_of_line/perform_paragraph_sets_ws_field
*> origin: languages/cobol/tests/cobol/test_perform_out_of_line.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 NAME PIC X(10) VALUE SPACES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    PERFORM SET-NAME.
    DISPLAY NAME.
    MOVE SPACES TO WS-VYBE-L
    STRING NAME DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "COBOL     "
        DISPLAY "FAIL: want [COBOL     ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.
SET-NAME.
    MOVE "COBOL" TO NAME.
    STOP RUN.

