*> vybe-test: cobol/pointers_arithmetic_ops/pointer_increment_external_call_compiles
*> origin: languages/cobol/tests/cobol/test_pointers_arithmetic_ops.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 P USAGE POINTER.
PROCEDURE DIVISION.
    CALL "PTR-INC" USING P.
    STOP RUN.

