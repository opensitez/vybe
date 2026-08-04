*> vybe-test: cobol/category_data_division/test_data_redefines_impact_group_move
*> origin: languages/cobol/tests/cobol/test_category_data_division.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. DATA-MOVE.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 GROUP.
          05 A PIC X(2) VALUE "HI".
          05 B PIC X(2) VALUE "JO".
       01 NUM REDEFINES GROUP PIC X(4).
       PROCEDURE DIVISION.
           MOVE GROUP TO NUM.
           DISPLAY NUM.
    MOVE SPACES TO WS-VYBE-L
    STRING NUM DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "HIJO"
        DISPLAY "FAIL: want [HIJO] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

