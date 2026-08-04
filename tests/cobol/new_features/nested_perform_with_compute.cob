*> vybe-test: cobol/new_features/nested_perform_with_compute
*> origin: languages/cobol/tests/cobol/test_new_features.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. NESTPERF.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TABLE.
   05 WS-ROW OCCURS 3 TIMES.
      10 WS-COL PIC 9(5) OCCURS 3 TIMES.
01 WS-I PIC 9(3).
01 WS-J PIC 9(3).
PROCEDURE DIVISION.
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > 3
        PERFORM VARYING WS-J FROM 1 BY 1 UNTIL WS-J > 3
            COMPUTE WS-COL(WS-J) = WS-I * WS-J
        END-PERFORM
    END-PERFORM.
    DISPLAY "Multiplication table done".
    STOP RUN.

