*> vybe-test: cobol/cobol2023_file_io/write_advancing_page
*> origin: languages/cobol/tests/cobol/test_cobol2023_file_io.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-REC PIC X(80) VALUE "Page header".
PROCEDURE DIVISION.
    WRITE WS-REC BEFORE ADVANCING PAGE.
    STOP RUN.

