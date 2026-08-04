*> vybe-test: cobol/async_threads/perform_until_with_counter
*> origin: languages/cobol/tests/cobol/test_async_threads.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 I PIC 9 VALUE 0.
PROCEDURE DIVISION.
    PERFORM UNTIL I >= 3
        ADD 1 TO I
        CALL "SUBE"
    END-PERFORM.
    STOP RUN.

