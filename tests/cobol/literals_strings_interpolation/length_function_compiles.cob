*> vybe-test: cobol/literals_strings_interpolation/length_function_compiles
*> origin: languages/cobol/tests/cobol/test_literals_strings_interpolation.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(10) VALUE "ABC".
01 L PIC 9(3).
PROCEDURE DIVISION.
    MOVE FUNCTION LENGTH(S) TO L.
    STOP RUN.

