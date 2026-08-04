*> vybe-test: cobol/cobol2023_string_ops/inspect_converting
*> origin: languages/cobol/tests/cobol/test_cobol2023_string_ops.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-STR PIC X(20) VALUE "Hello World".
PROCEDURE DIVISION.
    INSPECT WS-STR CONVERTING "abcdefghijklmnopqrstuvwxyz"
        TO "ABCDEFGHIJKLMNOPQRSTUVWXYZ".
    DISPLAY WS-STR.
    STOP RUN.

