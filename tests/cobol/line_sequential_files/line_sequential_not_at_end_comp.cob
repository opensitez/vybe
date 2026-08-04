*> vybe-test: cobol/line_sequential_files/line_sequential_not_at_end_compiles
*> origin: languages/cobol/tests/cobol/test_line_sequential_files.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT LF ASSIGN TO "l.txt" ORGANIZATION IS LINE SEQUENTIAL.
DATA DIVISION.
FILE SECTION.
FD LF.
01 LR PIC X(80).
PROCEDURE DIVISION.
    OPEN INPUT LF
    READ LF
        AT END DISPLAY "EOF"
        NOT AT END DISPLAY LR
    END-READ
    CLOSE LF.
    STOP RUN.

