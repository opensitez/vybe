*> vybe-test: cobol/intrinsics_extended/intrinsic_length_compiles
*> origin: languages/cobol/tests/cobol/test_intrinsics_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TXT PIC X(10) VALUE "HELLO".
01 WS-LEN PIC 9(3).
PROCEDURE DIVISION.
    MOVE FUNCTION LENGTH(WS-TXT) TO WS-LEN.
    STOP RUN.

