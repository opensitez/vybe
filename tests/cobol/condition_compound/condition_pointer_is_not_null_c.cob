*> vybe-test: cobol/condition_compound/condition_pointer_is_not_null_compiles
*> origin: languages/cobol/tests/cobol/test_condition_compound.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 P USAGE POINTER.
PROCEDURE DIVISION.
    IF P NOT = NULL
        DISPLAY "NOT NULL"
    END-IF.
    STOP RUN.

