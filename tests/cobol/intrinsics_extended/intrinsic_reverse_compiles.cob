*> vybe-test: cobol/intrinsics_extended/intrinsic_reverse_compiles
*> origin: languages/cobol/tests/cobol/test_intrinsics_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TXT PIC X(10) VALUE "HELLO".
01 WS-OUT PIC X(10).
PROCEDURE DIVISION.
    MOVE FUNCTION REVERSE(WS-TXT) TO WS-OUT.
    STOP RUN.

