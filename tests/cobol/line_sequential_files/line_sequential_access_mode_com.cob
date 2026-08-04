*> vybe-test: cobol/line_sequential_files/line_sequential_access_mode_compiles
*> origin: languages/cobol/tests/cobol/test_line_sequential_files.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT LF ASSIGN TO "l.txt" ORGANIZATION IS LINE SEQUENTIAL ACCESS MODE IS SEQUENTIAL.
DATA DIVISION.
FILE SECTION.
FD LF.
01 LR PIC X(80).
PROCEDURE DIVISION.
    STOP RUN.

