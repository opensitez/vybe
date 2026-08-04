*> vybe-test: cobol/cobol2023_file_io/rewrite_with_invalid_key
*> origin: languages/cobol/tests/cobol/test_cobol2023_file_io.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-REC PIC X(80) VALUE "Updated".
PROCEDURE DIVISION.
    REWRITE WS-REC
        INVALID KEY DISPLAY "Failed"
        NOT INVALID KEY DISPLAY "OK"
    END-REWRITE.
    STOP RUN.

