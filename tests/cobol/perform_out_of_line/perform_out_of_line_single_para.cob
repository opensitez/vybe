*> vybe-test: cobol/perform_out_of_line/perform_out_of_line_single_paragraph
*> origin: languages/cobol/tests/cobol/test_perform_out_of_line.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    PERFORM SHOW-MSG.
    STOP RUN.
SHOW-MSG.
    DISPLAY "HELLO FROM PARA".
    MOVE SPACES TO WS-VYBE-L
    STRING "HELLO FROM PARA" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "HELLO FROM PARA"
        DISPLAY "FAIL: want [HELLO FROM PARA] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

