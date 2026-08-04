*> vybe-test: cobol/arithmetic_control_flow/search_linear_with_at_end_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TBL.
   05 E OCCURS 3 TIMES INDEXED BY IDX.
      10 K PIC 9.
01 F PIC X VALUE "N".
PROCEDURE DIVISION.
    MOVE 1 TO K(1).
    MOVE 2 TO K(2).
    MOVE 3 TO K(3).
    SET IDX TO 1.
    SEARCH E
        AT END MOVE "N" TO F
        WHEN K(IDX) = 2 MOVE "Y" TO F
    END-SEARCH.
    DISPLAY F.
    STOP RUN.

