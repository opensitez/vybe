*> vybe-test: cobol/cobol2023_string_ops/unstring_tallying
*> origin: languages/cobol/tests/cobol/test_cobol2023_string_ops.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DATA PIC X(20) VALUE "A,B,C,D".
01 WS-P1 PIC X(5).
01 WS-P2 PIC X(5).
01 WS-P3 PIC X(5).
01 WS-P4 PIC X(5).
01 WS-COUNT PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    UNSTRING WS-DATA
        DELIMITED BY ","
        INTO WS-P1 WS-P2 WS-P3 WS-P4
        TALLYING IN WS-COUNT.
    DISPLAY WS-COUNT.
    STOP RUN.

