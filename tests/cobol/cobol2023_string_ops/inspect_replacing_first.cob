*> vybe-test: cobol/cobol2023_string_ops/inspect_replacing_first
*> origin: languages/cobol/tests/cobol/test_cobol2023_string_ops.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-STR PIC X(20) VALUE "AABAA".
PROCEDURE DIVISION.
    INSPECT WS-STR REPLACING FIRST "A" BY "X".
    DISPLAY WS-STR.
    STOP RUN.

