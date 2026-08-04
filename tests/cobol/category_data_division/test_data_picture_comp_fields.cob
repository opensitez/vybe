*> vybe-test: cobol/category_data_division/test_data_picture_comp_fields
*> origin: languages/cobol/tests/cobol/test_category_data_division.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. DATA-PIC.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 DECIMAL PIC S9(4)V99 VALUE -12.34.
       PROCEDURE DIVISION.
           DISPLAY DECIMAL.
    MOVE SPACES TO WS-VYBE-L
    STRING DECIMAL DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "-1234"
        DISPLAY "FAIL: want [-1234] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

