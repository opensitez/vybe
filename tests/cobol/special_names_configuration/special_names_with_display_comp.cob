*> vybe-test: cobol/special_names_configuration/special_names_with_display_compiles
*> origin: languages/cobol/tests/cobol/test_special_names_configuration.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    ALPHABET ALP IS ASCII.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC X(10) VALUE "ABC".
PROCEDURE DIVISION.
    DISPLAY X.
    STOP RUN.

