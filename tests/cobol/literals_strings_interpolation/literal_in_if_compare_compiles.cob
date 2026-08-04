*> vybe-test: cobol/literals_strings_interpolation/literal_in_if_compare_compiles
*> origin: languages/cobol/tests/cobol/test_literals_strings_interpolation.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC X(5) VALUE "YES".
PROCEDURE DIVISION.
    IF A = "YES" DISPLAY "OK" END-IF.
    STOP RUN.

