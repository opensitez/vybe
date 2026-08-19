*> vybe-test: cobol/category_data_editing/test_data_edit_complex_combined
*> origin: languages/cobol/tests/cobol/test_category_data_editing.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. EDIT-COMPLEX.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 NUM PIC S9(5)V99 VALUE -1234.56.
       01 EDITED PIC $$$,$$9.99DB.
       PROCEDURE DIVISION.
           MOVE NUM TO EDITED.
           DISPLAY "[" EDITED "]".
    MOVE SPACES TO WS-VYBE-L
    STRING "[" DELIMITED SIZE EDITED DELIMITED SIZE "]" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "[ $1,234.56DB]"
        DISPLAY "FAIL: want [[ $1,234.56DB]] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

