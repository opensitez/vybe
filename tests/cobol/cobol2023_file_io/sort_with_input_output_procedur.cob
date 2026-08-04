*> vybe-test: cobol/cobol2023_file_io/sort_with_input_output_procedure
*> origin: languages/cobol/tests/cobol/test_cobol2023_file_io.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-REC PIC X(80).
01 WS-KEY PIC X(10).
PROCEDURE DIVISION.
    SORT WS-SORT-FILE ON ASCENDING KEY WS-KEY
        USING WS-IN-FILE
        GIVING WS-OUT-FILE.
    STOP RUN.

