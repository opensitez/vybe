*> vybe-test: cobol/scope_terminators_nesting/scope_end_multiply_compiles
*> origin: languages/cobol/tests/cobol/test_scope_terminators_nesting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3) VALUE 2.
PROCEDURE DIVISION.
    MULTIPLY 3 BY A
    END-MULTIPLY.
    STOP RUN.

