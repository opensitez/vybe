*> vybe-test: cobol/pointers/pointer_null_comparison_in_if_compiles
*> origin: languages/cobol/tests/cobol/test_pointers.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 P USAGE POINTER.
PROCEDURE DIVISION.
    SET P TO NULL.
    IF P = NULL
        DISPLAY "NULL"
    END-IF.
    STOP RUN.

