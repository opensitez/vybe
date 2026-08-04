*> vybe-test: cobol/file_locking/close_with_lock_after_output_compiles
*> origin: languages/cobol/tests/cobol/test_file_locking.rs
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
PROCEDURE DIVISION.
    OPEN OUTPUT F.
    CLOSE F WITH LOCK.
    STOP RUN.

