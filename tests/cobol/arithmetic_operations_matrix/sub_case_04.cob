*> vybe-test: cobol/arithmetic_operations_matrix/sub_case_04
*> origin: languages/cobol/tests/cobol/test_arithmetic_operations_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 99 VALUE 9.
01 B PIC 9 VALUE 2.
PROCEDURE DIVISION.
    SUBTRACT B FROM A END-SUBTRACT.
    STOP RUN.

