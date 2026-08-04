*> vybe-test: cobol/end_of_file_handling/read_at_end_moves_flag_compiles
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
01 EOF-FLAG PIC X VALUE "N".
PROCEDURE DIVISION.
    OPEN INPUT F.
    READ F AT END MOVE "Y" TO EOF-FLAG END-READ.
    CLOSE F.
    STOP RUN.

