*> vybe-test: cobol/pointers_arithmetic_ops/pointer_callback_call_compiles
*> origin: languages/cobol/tests/cobol/test_pointers_arithmetic_ops.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 PP USAGE PROCEDURE-POINTER.
PROCEDURE DIVISION.
    CALL "INVOKE-CB" USING PP.
    STOP RUN.

