*> vybe-test: cobol/iterators_generators_control/iter_varying_by_step_compiles
*> origin: languages/cobol/tests/cobol/test_iterators_generators_control.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 I PIC 9 VALUE 1.
01 T PIC 9 OCCURS 5 TIMES.
PROCEDURE DIVISION.
    PERFORM VARYING I FROM 1 BY 2 UNTIL I > 5
        MOVE I TO T(I)
    END-PERFORM.
    STOP RUN.

