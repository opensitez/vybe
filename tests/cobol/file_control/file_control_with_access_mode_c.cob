*> vybe-test: cobol/file_control/file_control_with_access_mode_compiles
*> origin: languages/cobol/tests/cobol/test_file_control.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT F ASSIGN TO "f.dat"
        ORGANIZATION IS INDEXED
        ACCESS MODE IS DYNAMIC
        RECORD KEY IS K.
DATA DIVISION.
FILE SECTION.
FD F.
01 REC.
   05 K PIC 9(5).
PROCEDURE DIVISION.
    STOP RUN.

