*> vybe-test: cobol/open_close_modes/open_input_output_io_extend_compiles
*> origin: languages/cobol/tests/cobol/test_open_close_modes.rs
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
    OPEN INPUT F.
    CLOSE F.
    OPEN OUTPUT F.
    CLOSE F.
    OPEN I-O F.
    CLOSE F.
    OPEN EXTEND F.
    CLOSE F.
    STOP RUN.

