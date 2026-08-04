*> vybe-test: cobol/iterators_generators_control/iter_over_table_compiles
*> origin: languages/cobol/tests/cobol/test_iterators_generators_control.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T PIC X(2) OCCURS 3 TIMES.
01 I PIC 9 VALUE 1.
PROCEDURE DIVISION.
    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 3
        DISPLAY T(I)
    END-PERFORM.
    STOP RUN.

