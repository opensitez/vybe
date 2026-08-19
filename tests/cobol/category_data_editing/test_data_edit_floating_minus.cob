*> vybe-test: cobol/category_data_editing/test_data_edit_floating_minus
*> origin: languages/cobol/tests/cobol/test_category_data_editing.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. EDIT-FLOAT-MINUS.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 NUM PIC S9(3) VALUE -45.
       01 EDITED PIC ---,---.
       PROCEDURE DIVISION.
           MOVE NUM TO EDITED.
           DISPLAY "[" EDITED "]".
    MOVE SPACES TO WS-VYBE-L
    STRING "[" DELIMITED SIZE EDITED DELIMITED SIZE "]" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "[    -45]"
        DISPLAY "FAIL: want [[    -45]] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

