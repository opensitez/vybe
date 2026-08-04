*> vybe-test: cobol/file_sharing/open_input_and_output_same_file_sequence_compiles
*> origin: languages/cobol/tests/cobol/test_file_sharing.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT F1 ASSIGN TO "a.dat".
DATA DIVISION.
FILE SECTION.
FD F1.
01 R1 PIC X(10).
PROCEDURE DIVISION.
    OPEN OUTPUT F1.
    CLOSE F1.
    OPEN I-O F1.
    CLOSE F1.
    STOP RUN.

