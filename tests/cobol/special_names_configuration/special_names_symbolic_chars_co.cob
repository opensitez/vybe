*> vybe-test: cobol/special_names_configuration/special_names_symbolic_chars_compiles
*> origin: languages/cobol/tests/cobol/test_special_names_configuration.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    SYMBOLIC CHARACTERS CR IS 13.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC X.
PROCEDURE DIVISION.
    MOVE CR TO X.
    STOP RUN.

