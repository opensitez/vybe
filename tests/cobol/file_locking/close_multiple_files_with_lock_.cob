*> vybe-test: cobol/file_locking/close_multiple_files_with_lock_compiles
*> origin: languages/cobol/tests/cobol/test_file_locking.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT F1 ASSIGN TO "a.dat".
    SELECT F2 ASSIGN TO "b.dat".
DATA DIVISION.
FILE SECTION.
FD F1.
01 R1 PIC X(20).
FD F2.
01 R2 PIC X(20).
PROCEDURE DIVISION.
    OPEN I-O F1 F2.
    CLOSE F1 WITH LOCK F2 WITH LOCK.
    STOP RUN.

