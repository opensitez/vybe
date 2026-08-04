*> vybe-test: cobol/special_registers_detail/pointer_null_comparison_compiles
*> origin: languages/cobol/tests/cobol/test_special_registers_detail.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 P USAGE POINTER VALUE NULL.
PROCEDURE DIVISION.
    IF P = NULL
        DISPLAY "NULL"
    END-IF.
    STOP RUN.

