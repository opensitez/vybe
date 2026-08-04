*> vybe-test: cobol/string_functions_intrinsic/intrinsic_char_from_ord
*> origin: languages/cobol/tests/cobol/test_string_functions_intrinsic.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 C PIC X.
PROCEDURE DIVISION.
    MOVE FUNCTION CHAR(65) TO C.
    STOP RUN.

