*> vybe-test: cobol/scope_terminators_nesting/scope_end_subtract_not_size_error
*> origin: languages/cobol/tests/cobol/test_scope_terminators_nesting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3) VALUE 100.
PROCEDURE DIVISION.
    SUBTRACT 1 FROM A
    END-SUBTRACT.
    STOP RUN.

