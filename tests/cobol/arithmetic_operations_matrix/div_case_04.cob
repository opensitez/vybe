*> vybe-test: cobol/arithmetic_operations_matrix/div_case_04
*> origin: languages/cobol/tests/cobol/test_arithmetic_operations_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 99 VALUE 20.
01 B PIC 9 VALUE 3.
01 Q PIC 99.
01 M PIC 9.
PROCEDURE DIVISION.
    DIVIDE B INTO A GIVING Q REMAINDER M.
    STOP RUN.

