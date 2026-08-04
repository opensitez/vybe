*> vybe-test: cobol/intrinsics_string/test_intrinsics_substitute_case
*> origin: languages/cobol/tests/cobol/test_intrinsics_string.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-STR PIC X(10) VALUE "Hello World".
01 WS-DST PIC X(15).
PROCEDURE DIVISION.

    MOVE FUNCTION SUBSTITUTE-CASE(WS-STR "world" "COBOL") TO WS-DST.
    STOP RUN.

