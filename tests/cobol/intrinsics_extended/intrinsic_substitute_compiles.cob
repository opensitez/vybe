*> vybe-test: cobol/intrinsics_extended/intrinsic_substitute_compiles
*> origin: languages/cobol/tests/cobol/test_intrinsics_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TXT PIC X(20) VALUE "HELLO WORLD".
01 WS-OUT PIC X(20).
PROCEDURE DIVISION.
    MOVE FUNCTION SUBSTITUTE(WS-TXT "WORLD" "COBOL") TO WS-OUT.
    STOP RUN.

