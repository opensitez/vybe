*> vybe-test: cobol/cobol2023_string_ops/inspect_tallying_leading
*> origin: languages/cobol/tests/cobol/test_cobol2023_string_ops.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-STR PIC X(20) VALUE "AAABBA".
01 WS-COUNT PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    INSPECT WS-STR TALLYING WS-COUNT FOR LEADING "A".
    DISPLAY WS-COUNT.
    STOP RUN.

