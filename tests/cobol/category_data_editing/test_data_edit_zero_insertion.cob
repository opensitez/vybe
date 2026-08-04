*> vybe-test: cobol/category_data_editing/test_data_edit_zero_insertion
*> origin: languages/cobol/tests/cobol/test_category_data_editing.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. EDIT-ZERO.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 NUM PIC 9(4) VALUE 1234.
       01 EDITED PIC 990099.
       PROCEDURE DIVISION.
           MOVE NUM TO EDITED.
           DISPLAY EDITED.
    MOVE SPACES TO WS-VYBE-L
    STRING EDITED DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "120034"
        DISPLAY "FAIL: want [120034] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

