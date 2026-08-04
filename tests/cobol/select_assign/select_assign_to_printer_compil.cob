*> vybe-test: cobol/select_assign/select_assign_to_printer_compiles
*> origin: languages/cobol/tests/cobol/test_select_assign.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT LISTING ASSIGN TO PRINTER.
DATA DIVISION.
FILE SECTION.
FD LISTING.
01 R PIC X(80).
PROCEDURE DIVISION.
    STOP RUN.

