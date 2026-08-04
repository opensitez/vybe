*> vybe-test: cobol/iterators_generators_control/iter_nested_table_traversal_compiles
*> origin: languages/cobol/tests/cobol/test_iterators_generators_control.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 I PIC 9 VALUE 1.
01 J PIC 9 VALUE 1.
01 T.
   05 R OCCURS 2 TIMES.
      10 C PIC 9 OCCURS 2 TIMES.
PROCEDURE DIVISION.
    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 2
        PERFORM VARYING J FROM 1 BY 1 UNTIL J > 2
            MOVE J TO C(I J)
        END-PERFORM
    END-PERFORM.
    STOP RUN.

