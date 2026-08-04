*> vybe-test: cobol/file_control/file_control_assign_to_dynamic_field_compiles
*> origin: languages/cobol/tests/cobol/test_file_control.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT F ASSIGN TO WS-NAME.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NAME PIC X(20) VALUE "f.dat".
FILE SECTION.
FD F.
01 R PIC X(20).
PROCEDURE DIVISION.
    STOP RUN.

