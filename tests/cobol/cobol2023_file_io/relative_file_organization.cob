*> vybe-test: cobol/cobol2023_file_io/relative_file_organization
*> origin: languages/cobol/tests/cobol/test_cobol2023_file_io.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT REL-FILE ASSIGN TO "relative.dat"
        ORGANIZATION IS RELATIVE
        ACCESS MODE IS RANDOM.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DUMMY PIC X(1).
PROCEDURE DIVISION.
    DISPLAY "Relative file defined".
    STOP RUN.

