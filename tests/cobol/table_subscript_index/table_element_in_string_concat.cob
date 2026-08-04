*> vybe-test: cobol/table_subscript_index/table_element_in_string_concat
*> origin: languages/cobol/tests/cobol/test_table_subscript_index.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 NAMES.
   05 NAME PIC X(5) OCCURS 3 TIMES.
01 RESULT PIC X(20).
PROCEDURE DIVISION.
    MOVE "ALICE" TO NAME(1).
    MOVE "BOB  " TO NAME(2).
    STRING NAME(1) DELIMITED BY SPACE " " DELIMITED BY SIZE
           NAME(2) DELIMITED BY SPACE INTO RESULT.
    STOP RUN.

