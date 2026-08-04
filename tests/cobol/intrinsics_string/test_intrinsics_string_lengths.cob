*> vybe-test: cobol/intrinsics_string/test_intrinsics_string_lengths
*> origin: languages/cobol/tests/cobol/test_intrinsics_string.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-STR PIC X(10) VALUE "hello".
01 WS-LEN PIC 9(5).
PROCEDURE DIVISION.

    COMPUTE WS-LEN = FUNCTION BYTE-LENGTH(WS-STR).
    COMPUTE WS-LEN = FUNCTION STORED-CHAR-LENGTH(WS-STR).
    STOP RUN.

