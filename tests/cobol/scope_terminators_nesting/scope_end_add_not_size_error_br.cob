*> vybe-test: cobol/scope_terminators_nesting/scope_end_add_not_size_error_branch
*> origin: languages/cobol/tests/cobol/test_scope_terminators_nesting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3) VALUE 5.
PROCEDURE DIVISION.
    ADD 10 TO A
    NOT ON SIZE ERROR
        DISPLAY "OK"
    END-ADD.
    STOP RUN.

