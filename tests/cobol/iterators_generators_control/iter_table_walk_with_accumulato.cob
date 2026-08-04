*> vybe-test: cobol/iterators_generators_control/iter_table_walk_with_accumulator_compiles
*> origin: languages/cobol/tests/cobol/test_iterators_generators_control.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T PIC 9 OCCURS 3 TIMES.
01 I PIC 9 VALUE 1.
01 S PIC 99 VALUE 0.
PROCEDURE DIVISION.
    MOVE 1 TO T(1).
    MOVE 2 TO T(2).
    MOVE 3 TO T(3).
    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 3
        ADD T(I) TO S
    END-PERFORM.
    DISPLAY S.
    STOP RUN.

