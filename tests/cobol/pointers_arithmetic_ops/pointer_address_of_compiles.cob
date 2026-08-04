*> vybe-test: cobol/pointers_arithmetic_ops/pointer_address_of_compiles
*> origin: languages/cobol/tests/cobol/test_pointers_arithmetic_ops.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 P USAGE POINTER.
01 A PIC X(5).
PROCEDURE DIVISION.
    SET P TO ADDRESS OF A.
    STOP RUN.

