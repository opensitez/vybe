*> vybe-test: cobol/final_features/command_line_program
*> origin: languages/cobol/tests/cobol/test_final_features.rs
*> vybe-test-mode: compile

IDENTIFICATION DIVISION.
PROGRAM-ID. CLIAPP.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-ARGS PIC X(200).
01 WS-NAME PIC X(50).
PROCEDURE DIVISION.
    ACCEPT WS-ARGS FROM COMMAND-LINE.
    DISPLAY "Arguments: " WS-ARGS.
    ACCEPT WS-NAME.
    DISPLAY "Hello " WS-NAME.
    STOP RUN.

