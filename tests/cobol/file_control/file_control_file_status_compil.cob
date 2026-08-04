*> vybe-test: cobol/file_control/file_control_file_status_compiles
*> origin: languages/cobol/tests/cobol/test_file_control.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT F ASSIGN TO "f.dat" FILE STATUS IS FS.
DATA DIVISION.
FILE SECTION.
FD F.
01 R PIC X(20).
WORKING-STORAGE SECTION.
01 FS PIC XX.
PROCEDURE DIVISION.
    STOP RUN.

