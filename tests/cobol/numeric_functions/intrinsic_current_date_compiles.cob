*> vybe-test: cobol/numeric_functions/intrinsic_current_date_compiles
*> origin: languages/cobol/tests/cobol/test_numeric_functions.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TODAY PIC X(21).
PROCEDURE DIVISION.
    MOVE FUNCTION CURRENT-DATE TO TODAY.
    STOP RUN.

