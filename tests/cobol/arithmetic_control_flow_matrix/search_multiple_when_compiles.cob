*> vybe-test: cobol/arithmetic_control_flow_matrix/search_multiple_when_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TAB.
   05 E OCCURS 4 TIMES INDEXED BY I.
      10 K PIC 9.
PROCEDURE DIVISION.
    MOVE 1 TO K(1).
    MOVE 2 TO K(2).
    MOVE 3 TO K(3).
    MOVE 4 TO K(4).
    SET I TO 1.
    SEARCH E
        WHEN K(I) = 1 DISPLAY "A"
        WHEN K(I) = 4 DISPLAY "D"
    END-SEARCH.
    STOP RUN.

