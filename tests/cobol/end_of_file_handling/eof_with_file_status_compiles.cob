*> vybe-test: cobol/end_of_file_handling/eof_with_file_status_compiles
*> origin: languages/cobol/tests/cobol/test_end_of_file_handling.rs
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
    OPEN INPUT F.
    READ F AT END DISPLAY FS END-READ.
    CLOSE F.
    STOP RUN.

