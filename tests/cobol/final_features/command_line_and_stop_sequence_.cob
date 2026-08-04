*> vybe-test: cobol/final_features/command_line_and_stop_sequence_program
*> origin: languages/cobol/tests/cobol/test_final_features.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. CMD2.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-ARGS PIC X(80).
01 WS-COUNT PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    ACCEPT WS-ARGS FROM COMMAND-LINE.
    MOVE 1 TO WS-COUNT.
    IF WS-COUNT = 1
        DISPLAY "CMD READY"
    END-IF.
    STOP RUN.

