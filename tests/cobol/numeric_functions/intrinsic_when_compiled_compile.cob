*> vybe-test: cobol/numeric_functions/intrinsic_when_compiled_compiles
*> origin: languages/cobol/tests/cobol/test_numeric_functions.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 COMPILED-DATE PIC X(21).
PROCEDURE DIVISION.
    MOVE FUNCTION WHEN-COMPILED TO COMPILED-DATE.
    STOP RUN.

