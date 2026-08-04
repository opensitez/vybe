*> vybe-test: cobol/cobol2023_file_io/delete_with_invalid_key
*> origin: languages/cobol/tests/cobol/test_cobol2023_file_io.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-REC PIC X(80).
PROCEDURE DIVISION.
    DELETE WS-FILE RECORD
        INVALID KEY DISPLAY "Not found"
        NOT INVALID KEY DISPLAY "Deleted"
    END-DELETE.
    STOP RUN.

