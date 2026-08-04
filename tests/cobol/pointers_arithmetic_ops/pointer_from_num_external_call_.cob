*> vybe-test: cobol/pointers_arithmetic_ops/pointer_from_num_external_call_compiles
*> origin: languages/cobol/tests/cobol/test_pointers_arithmetic_ops.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 P USAGE POINTER.
01 N PIC 9(10) VALUE 100.
PROCEDURE DIVISION.
    CALL "NUM-TO-PTR" USING N P.
    STOP RUN.

