*> vybe-test: cobol/category_data_editing/test_data_edit_insertion_characters
*> origin: languages/cobol/tests/cobol/test_category_data_editing.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. EDIT-INS.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 NUM PIC 9(6) VALUE 123456.
       01 EDITED PIC 99/99/99.
       PROCEDURE DIVISION.
           MOVE NUM TO EDITED.
           DISPLAY EDITED.
    MOVE SPACES TO WS-VYBE-L
    STRING EDITED DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "12/34/56"
        DISPLAY "FAIL: want [12/34/56] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

