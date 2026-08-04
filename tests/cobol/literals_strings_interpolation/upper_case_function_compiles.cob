*> vybe-test: cobol/literals_strings_interpolation/upper_case_function_compiles
*> origin: languages/cobol/tests/cobol/test_literals_strings_interpolation.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(10) VALUE "abc".
01 O PIC X(10).
PROCEDURE DIVISION.
    MOVE FUNCTION UPPER-CASE(S) TO O.
    STOP RUN.

