*> vybe-test: cobol/data_group_level/data_group_nested_with_occurs
*> origin: languages/cobol/tests/cobol/test_data_group_level.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 OUTER.
   05 INNER PIC 9(2) OCCURS 5 TIMES.
PROCEDURE DIVISION.
    MOVE 1 TO INNER(1).
    MOVE 2 TO INNER(2).
    STOP RUN.

