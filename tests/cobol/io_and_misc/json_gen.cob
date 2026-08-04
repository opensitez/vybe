*> vybe-test: cobol/io_and_misc/json_gen
*> origin: languages/cobol/tests/cobol/test_io_and_misc.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 REC.
   05 NAME PIC X(10) VALUE "Alice".
   05 AGE PIC 9(3) VALUE 30.
01 J PIC X(100).
PROCEDURE DIVISION.
    JSON GENERATE J FROM REC.
    STOP RUN.

