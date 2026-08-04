*> vybe-test: cobol/value_all_forms/value_quotes_single_char
*> origin: languages/cobol/tests/cobol/test_value_all_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X VALUE QUOTE.
PROCEDURE DIVISION.
    DISPLAY S.
    STOP RUN.

