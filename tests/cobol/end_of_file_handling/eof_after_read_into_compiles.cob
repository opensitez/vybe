*> vybe-test: cobol/end_of_file_handling/eof_after_read_into_compiles
*> origin: languages/cobol/tests/cobol/test_end_of_file_handling.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT F ASSIGN TO "f.dat".
DATA DIVISION.
FILE SECTION.
FD F.
01 R PIC X(20).
WORKING-STORAGE SECTION.
01 W PIC X(20).
PROCEDURE DIVISION.
    OPEN INPUT F.
    READ F INTO W AT END DISPLAY "EOF" END-READ.
    CLOSE F.
    STOP RUN.

