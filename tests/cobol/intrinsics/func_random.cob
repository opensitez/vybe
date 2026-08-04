*> vybe-test: cobol/intrinsics/func_random
*> origin: languages/cobol/tests/cobol/test_intrinsics.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 R PIC 9(10) VALUE 0.
01 A PIC 9(10) VALUE 10.
01 B PIC 9(10) VALUE 20.
01 C PIC 9(10) VALUE 30.
PROCEDURE DIVISION.
    COMPUTE R = FUNCTION RANDOM.
    STOP RUN.

