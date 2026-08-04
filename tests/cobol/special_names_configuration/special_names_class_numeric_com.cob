*> vybe-test: cobol/special_names_configuration/special_names_class_numeric_compiles
*> origin: languages/cobol/tests/cobol/test_special_names_configuration.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    CLASS DIGITS IS "0" THRU "9".
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC X(1).
PROCEDURE DIVISION.
    IF X IS DIGITS DISPLAY "D" END-IF.
    STOP RUN.

