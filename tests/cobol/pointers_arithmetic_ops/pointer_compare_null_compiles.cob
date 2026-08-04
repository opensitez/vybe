*> vybe-test: cobol/pointers_arithmetic_ops/pointer_compare_null_compiles
*> origin: languages/cobol/tests/cobol/test_pointers_arithmetic_ops.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 P USAGE POINTER.
PROCEDURE DIVISION.
    IF P = NULL DISPLAY "N" END-IF.
    STOP RUN.

