*> vybe-test: cobol/intrinsics_extended/intrinsic_upper_case_compiles
*> origin: languages/cobol/tests/cobol/test_intrinsics_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TXT PIC X(10) VALUE "abc".
01 WS-OUT PIC X(10).
PROCEDURE DIVISION.
    MOVE FUNCTION UPPER-CASE(WS-TXT) TO WS-OUT.
    STOP RUN.

