*> vybe-test: cobol/category_data_division/test_data_value_hex
*> origin: languages/cobol/tests/cobol/test_category_data_division.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. DATA-HEX.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 VAL PIC X VALUE X"41".
       PROCEDURE DIVISION.
           DISPLAY VAL.
    MOVE SPACES TO WS-VYBE-L
    STRING VAL DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "A"
        DISPLAY "FAIL: want [A] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

