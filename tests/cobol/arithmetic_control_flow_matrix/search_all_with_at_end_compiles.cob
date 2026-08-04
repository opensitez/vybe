*> vybe-test: cobol/arithmetic_control_flow_matrix/search_all_with_at_end_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TAB.
   05 E OCCURS 4 TIMES ASCENDING KEY IS K INDEXED BY I.
      10 K PIC 9(2).
01 F PIC X VALUE "N".
PROCEDURE DIVISION.
    MOVE 1 TO K(1).
    MOVE 2 TO K(2).
    MOVE 3 TO K(3).
    MOVE 4 TO K(4).
    SEARCH ALL E
        AT END MOVE "N" TO F
        WHEN K(I) = 3 MOVE "Y" TO F
    END-SEARCH.
    DISPLAY F.
    STOP RUN.

