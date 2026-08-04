*> vybe-test: cobol/arithmetic_operations_matrix/add_case_07
*> origin: languages/cobol/tests/cobol/test_arithmetic_operations_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 9.
01 B PIC 9 VALUE 9.
PROCEDURE DIVISION.
    ADD A TO B ON SIZE ERROR DISPLAY "E" END-ADD.
    STOP RUN.

