*> vybe-test: cobol/cobol2023_string_ops/reference_modification_no_length
*> origin: languages/cobol/tests/cobol/test_cobol2023_string_ops.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-STR PIC X(20) VALUE "Hello World".
01 WS-RESULT PIC X(20).
PROCEDURE DIVISION.
    MOVE WS-STR(7:) TO WS-RESULT.
    DISPLAY WS-RESULT.
    STOP RUN.

