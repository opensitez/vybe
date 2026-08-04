*> vybe-test: cobol/pointers_arithmetic_ops/pointer_diff_external_call_compiles
*> origin: languages/cobol/tests/cobol/test_pointers_arithmetic_ops.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 P1 USAGE POINTER.
01 P2 USAGE POINTER.
01 D PIC 9(5).
PROCEDURE DIVISION.
    CALL "PTR-DIFF" USING P1 P2 D.
    STOP RUN.

