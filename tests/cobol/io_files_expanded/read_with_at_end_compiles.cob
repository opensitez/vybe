*> vybe-test: cobol/io_files_expanded/read_with_at_end_compiles
*> origin: languages/cobol/tests/cobol/test_io_files_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    READ WS-FILE
        AT END DISPLAY "EOF"
    END-READ.
    STOP RUN.

