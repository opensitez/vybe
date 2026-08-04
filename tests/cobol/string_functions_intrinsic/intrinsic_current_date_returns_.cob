*> vybe-test: cobol/string_functions_intrinsic/intrinsic_current_date_returns_21_chars
*> origin: languages/cobol/tests/cobol/test_string_functions_intrinsic.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TODAY PIC X(21).
PROCEDURE DIVISION.
    MOVE FUNCTION CURRENT-DATE TO TODAY.
    DISPLAY TODAY.
    STOP RUN.

