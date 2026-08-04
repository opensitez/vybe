*> vybe-test: cobol/category_data_division/test_data_redefines
*> origin: languages/cobol/tests/cobol/test_category_data_division.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. DATA-REDEF.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 DATA-ITEM PIC X(4) VALUE "1234".
       01 DATA-NUM REDEFINES DATA-ITEM PIC 9(4).
       PROCEDURE DIVISION.
           DISPLAY DATA-NUM.
    MOVE SPACES TO WS-VYBE-L
    STRING DATA-NUM DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "1234"
        DISPLAY "FAIL: want [1234] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

