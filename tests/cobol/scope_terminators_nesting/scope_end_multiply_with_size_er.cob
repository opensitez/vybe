*> vybe-test: cobol/scope_terminators_nesting/scope_end_multiply_with_size_error
*> origin: languages/cobol/tests/cobol/test_scope_terminators_nesting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3) VALUE 999.
PROCEDURE DIVISION.
    MULTIPLY 999 BY A
    ON SIZE ERROR
        DISPLAY "OVERFLOW"
    END-MULTIPLY.
    STOP RUN.

