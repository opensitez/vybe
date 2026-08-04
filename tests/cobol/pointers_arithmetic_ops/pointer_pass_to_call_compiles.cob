*> vybe-test: cobol/pointers_arithmetic_ops/pointer_pass_to_call_compiles
*> origin: languages/cobol/tests/cobol/test_pointers_arithmetic_ops.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 P USAGE POINTER.
PROCEDURE DIVISION.
    CALL "PTR-USE" USING P.
    STOP RUN.

