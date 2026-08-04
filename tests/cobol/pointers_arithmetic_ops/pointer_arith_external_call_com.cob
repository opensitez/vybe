*> vybe-test: cobol/pointers_arithmetic_ops/pointer_arith_external_call_compiles
*> origin: languages/cobol/tests/cobol/test_pointers_arithmetic_ops.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 P USAGE POINTER.
01 N PIC 9(3) VALUE 1.
PROCEDURE DIVISION.
    CALL "PTR-ADD" USING P N.
    STOP RUN.

