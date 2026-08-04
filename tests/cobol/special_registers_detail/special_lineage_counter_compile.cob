*> vybe-test: cobol/special_registers_detail/special_lineage_counter_compiles
*> origin: languages/cobol/tests/cobol/test_special_registers_detail.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT REPORT-FILE ASSIGN TO "output.txt".
DATA DIVISION.
FILE SECTION.
FD REPORT-FILE LINAGE IS 60 LINES.
01 REPORT-REC PIC X(80).
WORKING-STORAGE SECTION.
PROCEDURE DIVISION.
    STOP RUN.

