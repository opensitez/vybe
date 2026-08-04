*> vybe-test: cobol/arithmetic_operations_matrix/div_case_07
*> origin: languages/cobol/tests/cobol/test_arithmetic_operations_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 99 VALUE 12.
01 B PIC 9 VALUE 3.
PROCEDURE DIVISION.
    DIVIDE B INTO A NOT ON SIZE ERROR DISPLAY "O" END-DIVIDE.
    STOP RUN.

