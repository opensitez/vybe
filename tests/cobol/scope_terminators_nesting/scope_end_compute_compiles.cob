*> vybe-test: cobol/scope_terminators_nesting/scope_end_compute_compiles
*> origin: languages/cobol/tests/cobol/test_scope_terminators_nesting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 R PIC 9(4) VALUE 0.
PROCEDURE DIVISION.
    COMPUTE R = 2 + 2
    END-COMPUTE.
    STOP RUN.

