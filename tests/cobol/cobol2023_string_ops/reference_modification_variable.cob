*> vybe-test: cobol/cobol2023_string_ops/reference_modification_variable
*> origin: languages/cobol/tests/cobol/test_cobol2023_string_ops.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-STR PIC X(20) VALUE "Hello World".
01 WS-POS PIC 9(3) VALUE 7.
01 WS-LEN PIC 9(3) VALUE 5.
01 WS-RESULT PIC X(10).
PROCEDURE DIVISION.
    MOVE WS-STR(WS-POS:WS-LEN) TO WS-RESULT.
    DISPLAY WS-RESULT.
    STOP RUN.

