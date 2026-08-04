*> vybe-test: cobol/cobol2023_intrinsics/substitute_multiple
*> origin: languages/cobol/tests/cobol/test_cobol2023_intrinsics.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-STR PIC X(30) VALUE "Hello World".
01 WS-RESULT PIC X(30).
PROCEDURE DIVISION.
    MOVE FUNCTION SUBSTITUTE(WS-STR "Hello" "Goodbye")
        TO WS-RESULT.
    DISPLAY WS-RESULT.
    STOP RUN.

