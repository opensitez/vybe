*> vybe-test: cobol/intrinsics_bit/test_intrinsics_hex_conversions
*> origin: languages/cobol/tests/cobol/test_intrinsics_bit.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-HEX PIC X(2).
01 WS-CHAR PIC X.
PROCEDURE DIVISION.

    MOVE FUNCTION HEX-OF("A") TO WS-HEX.
    MOVE FUNCTION HEX-TO-CHAR("41") TO WS-CHAR.
    STOP RUN.

