*> vybe-test: cobol/cobol2023_file_io/write_advancing_lines
*> origin: languages/cobol/tests/cobol/test_cobol2023_file_io.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-REC PIC X(80) VALUE "Test record".
PROCEDURE DIVISION.
    WRITE WS-REC AFTER ADVANCING 2 LINES.
    STOP RUN.

