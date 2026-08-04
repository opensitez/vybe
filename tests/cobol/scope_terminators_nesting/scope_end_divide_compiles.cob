*> vybe-test: cobol/scope_terminators_nesting/scope_end_divide_compiles
*> origin: languages/cobol/tests/cobol/test_scope_terminators_nesting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3) VALUE 10.
01 R PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    DIVIDE 2 INTO A GIVING R
    END-DIVIDE.
    STOP RUN.

