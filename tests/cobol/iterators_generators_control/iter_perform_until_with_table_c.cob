*> vybe-test: cobol/iterators_generators_control/iter_perform_until_with_table_condition_compiles
*> origin: languages/cobol/tests/cobol/test_iterators_generators_control.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 I PIC 9 VALUE 1.
01 T PIC 9 OCCURS 3 TIMES.
01 F PIC 9 VALUE 0.
PROCEDURE DIVISION.
    MOVE 1 TO T(1). MOVE 2 TO T(2). MOVE 3 TO T(3).
    PERFORM UNTIL I > 3 OR F = 1
        IF T(I) = 2 MOVE 1 TO F END-IF
        ADD 1 TO I
    END-PERFORM.
    STOP RUN.

