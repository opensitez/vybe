*> vybe-test: cobol/final_features/accept_command_line
*> origin: languages/cobol/tests/cobol/test_final_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-ARGS PIC X(100).
PROCEDURE DIVISION.
    ACCEPT WS-ARGS FROM COMMAND-LINE.
    STOP RUN.

