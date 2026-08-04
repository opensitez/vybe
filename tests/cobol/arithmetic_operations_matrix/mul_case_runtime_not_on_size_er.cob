*> vybe-test: cobol/arithmetic_operations_matrix/mul_case_runtime_not_on_size_error_path
*> origin: languages/cobol/tests/cobol/test_arithmetic_operations_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 3.
01 B PIC 9 VALUE 2.
PROCEDURE DIVISION.
    MULTIPLY A BY B NOT ON SIZE ERROR DISPLAY "OK" END-MULTIPLY.
    STOP RUN.

