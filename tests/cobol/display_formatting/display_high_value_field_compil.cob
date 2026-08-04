*> vybe-test: cobol/display_formatting/display_high_value_field_compiles
*> origin: languages/cobol/tests/cobol/test_display_formatting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(4).
PROCEDURE DIVISION.
    MOVE HIGH-VALUES TO S.
    DISPLAY S.
    STOP RUN.

