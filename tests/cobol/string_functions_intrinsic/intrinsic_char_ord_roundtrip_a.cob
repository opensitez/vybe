*> vybe-test: cobol/string_functions_intrinsic/intrinsic_char_ord_roundtrip_a
*> origin: languages/cobol/tests/cobol/test_string_functions_intrinsic.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 C PIC X.
01 N PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    MOVE FUNCTION CHAR(65) TO C.
    COMPUTE N = FUNCTION ORD(C).
    STOP RUN.

