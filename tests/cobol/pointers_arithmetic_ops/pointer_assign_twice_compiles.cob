*> vybe-test: cobol/pointers_arithmetic_ops/pointer_assign_twice_compiles
*> origin: languages/cobol/tests/cobol/test_pointers_arithmetic_ops.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 P USAGE POINTER.
01 A PIC X(5).
01 B PIC X(5).
PROCEDURE DIVISION.
    SET P TO ADDRESS OF A.
    SET P TO ADDRESS OF B.
    STOP RUN.

