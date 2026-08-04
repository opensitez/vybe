*> vybe-test: cobol/cobol2023_file_io/open_multiple_modes
*> origin: languages/cobol/tests/cobol/test_cobol2023_file_io.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DUMMY PIC X(1).
PROCEDURE DIVISION.
    OPEN INPUT WS-IN-FILE
         OUTPUT WS-OUT-FILE.
    CLOSE WS-IN-FILE WS-OUT-FILE.
    STOP RUN.

