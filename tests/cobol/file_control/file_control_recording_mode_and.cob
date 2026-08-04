*> vybe-test: cobol/file_control/file_control_recording_mode_and_block_compiles
*> origin: languages/cobol/tests/cobol/test_file_control.rs
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
    STOP RUN.

