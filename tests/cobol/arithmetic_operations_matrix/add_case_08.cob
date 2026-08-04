*> vybe-test: cobol/arithmetic_operations_matrix/add_case_08
*> origin: languages/cobol/tests/cobol/test_arithmetic_operations_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 2.
01 B PIC 9 VALUE 2.
PROCEDURE DIVISION.
    ADD A TO B NOT ON SIZE ERROR DISPLAY "O" END-ADD.
    STOP RUN.

