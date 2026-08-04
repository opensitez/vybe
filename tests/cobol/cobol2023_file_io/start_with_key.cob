*> vybe-test: cobol/cobol2023_file_io/start_with_key
*> origin: languages/cobol/tests/cobol/test_cobol2023_file_io.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-KEY PIC X(10) VALUE "ABC".
PROCEDURE DIVISION.
    START WS-FILE KEY >= WS-KEY
        INVALID KEY DISPLAY "Not found"
        NOT INVALID KEY DISPLAY "Found"
    END-START.
    STOP RUN.

