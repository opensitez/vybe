*> vybe-test: cobol/file_status_codes/file_status_open_io_compiles
*> origin: languages/cobol/tests/cobol/test_file_status_codes.rs
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
    OPEN I-O F.
    CLOSE F.
    STOP RUN.

