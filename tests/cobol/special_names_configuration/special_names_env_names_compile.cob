*> vybe-test: cobol/special_names_configuration/special_names_env_names_compiles
*> origin: languages/cobol/tests/cobol/test_special_names_configuration.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    CONSOLE IS CRT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC X(20).
PROCEDURE DIVISION.
    DISPLAY X UPON CRT.
    STOP RUN.

