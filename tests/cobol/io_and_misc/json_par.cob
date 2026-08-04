*> vybe-test: cobol/io_and_misc/json_par
*> origin: languages/cobol/tests/cobol/test_io_and_misc.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 J PIC X(100) VALUE '{"name":"Bob"}'.
01 REC.
   05 NAME PIC X(10).
PROCEDURE DIVISION.
    JSON PARSE J INTO REC.
    STOP RUN.

