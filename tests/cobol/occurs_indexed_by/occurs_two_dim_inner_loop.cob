*> vybe-test: cobol/occurs_indexed_by/occurs_two_dim_inner_loop
*> origin: languages/cobol/tests/cobol/test_occurs_indexed_by.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 MATRIX.
   05 ROW OCCURS 3 TIMES INDEXED BY RI.
      10 COL PIC 9 OCCURS 3 TIMES INDEXED BY CI.
PROCEDURE DIVISION.
    PERFORM VARYING RI FROM 1 BY 1 UNTIL RI > 3
        PERFORM VARYING CI FROM 1 BY 1 UNTIL CI > 3
            MOVE 0 TO COL(RI CI)
        END-PERFORM
    END-PERFORM.
    STOP RUN.

