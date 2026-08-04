*> vybe-test: cobol/special_names_configuration/special_names_combined_compiles
*> origin: languages/cobol/tests/cobol/test_special_names_configuration.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    DECIMAL-POINT IS COMMA
    ALPHABET A1 IS STANDARD-1.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC X(5).
PROCEDURE DIVISION.
    DISPLAY X.
    STOP RUN.

