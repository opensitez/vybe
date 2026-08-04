*> vybe-test: cobol/scope_terminators_nesting/scope_end_compute_with_on_size_error
*> origin: languages/cobol/tests/cobol/test_scope_terminators_nesting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 R PIC 9(2) VALUE 0.
PROCEDURE DIVISION.
    COMPUTE R = 99 * 99
    ON SIZE ERROR
        DISPLAY "TOO BIG"
    END-COMPUTE.
    STOP RUN.

