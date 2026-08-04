*> vybe-test: cobol/arithmetic_control_flow_matrix/search_linear_basic_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TAB.
   05 E OCCURS 3 TIMES INDEXED BY I.
      10 K PIC 9.
PROCEDURE DIVISION.
    MOVE 1 TO K(1).
    MOVE 2 TO K(2).
    MOVE 3 TO K(3).
    SET I TO 1.
    SEARCH E
        WHEN K(I) = 2 DISPLAY "Y"
    END-SEARCH.
    STOP RUN.

